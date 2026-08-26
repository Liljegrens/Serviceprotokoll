from flask import Flask, jsonify, request, send_from_directory, session
import sqlite3, json, os, secrets, uuid
from datetime import datetime
import openpyxl

BASE    = os.path.dirname(os.path.abspath(__file__))
DB      = os.path.join(BASE, 'serviceprotokoll.db')
UPLOADS = os.path.join(BASE, 'uploads')
os.makedirs(UPLOADS, exist_ok=True)
app  = Flask(__name__, static_folder=BASE)
_key_file = os.path.join(BASE, '.secret_key')
if os.environ.get('SECRET_KEY'):
    app.secret_key = os.environ['SECRET_KEY']
elif os.path.exists(_key_file):
    app.secret_key = open(_key_file).read().strip()
else:
    app.secret_key = secrets.token_hex(32)
    open(_key_file, 'w').write(app.secret_key)

# ── Database ────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as db:
        db.executescript("""
        CREATE TABLE IF NOT EXISTS machines (
            nr         TEXT PRIMARY KEY,
            kund       TEXT DEFAULT '',
            anlaggning TEXT DEFAULT '',
            fabrikat   TEXT DEFAULT '',
            modell     TEXT DEFAULT '',
            inkopar    TEXT DEFAULT '',
            notering   TEXT DEFAULT '',
            adress     TEXT DEFAULT '',
            stad       TEXT DEFAULT '',
            kontakt    TEXT DEFAULT '',
            telefon    TEXT DEFAULT '',
            tillvnr    TEXT DEFAULT '',
            arsmodell  TEXT DEFAULT '',
            beskrivning TEXT DEFAULT '',
            agare      TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS protocols (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            maskin_nr        TEXT NOT NULL,
            datum            TEXT,
            type             TEXT,
            kund             TEXT,
            anlaggning       TEXT,
            tekniker         TEXT,
            modell           TEXT,
            items_json       TEXT,
            saved_at         TEXT,
            resolved_json    TEXT DEFAULT '{}',
            workflow_status  TEXT DEFAULT 'inkommen',
            workflow_log     TEXT DEFAULT '[]',
            machine_type     TEXT DEFAULT ''
        );
        CREATE INDEX IF NOT EXISTS idx_protocols_maskin ON protocols(maskin_nr);
        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            name     TEXT NOT NULL,
            pin      TEXT NOT NULL,
            role     TEXT DEFAULT 'tekniker'
        );
        """)

init_db()

def seed_users():
    default_users = os.environ.get('DEFAULT_USERS', '')
    if not default_users:
        return
    with get_db() as db:
        count = db.execute('SELECT COUNT(*) FROM users').fetchone()[0]
        if count > 0:
            return
        for entry in default_users.split(','):
            parts = entry.strip().split(':')
            if len(parts) == 3:
                name, pin, role = parts
                db.execute('INSERT INTO users (name, pin, role) VALUES (?,?,?)', (name.strip(), pin.strip(), role.strip()))

seed_users()

# ── Serve frontend ───────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(BASE, 'serviceprotokoll.html')

@app.route('/manifest.json')
def manifest():
    return send_from_directory(BASE, 'manifest.json')

@app.route('/logo.png')
def logo():
    return send_from_directory(BASE, 'logo.png')

# ── Users & auth ─────────────────────────────────────────────

@app.route('/api/users', methods=['GET'])
def list_users():
    with get_db() as db:
        rows = db.execute('SELECT id, name, role FROM users ORDER BY name').fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/users', methods=['POST'])
def create_user():
    data = request.get_json()
    name = (data.get('name') or '').strip()
    pin  = (data.get('pin') or '').strip()
    role = data.get('role', 'tekniker')
    if not name or not pin:
        return jsonify({'error': 'Namn och PIN krävs'}), 400
    with get_db() as db:
        db.execute('INSERT INTO users (name, pin, role) VALUES (?,?,?)', (name, pin, role))
    return jsonify({'ok': True}), 201

@app.route('/api/users/<int:uid>', methods=['DELETE'])
def delete_user(uid):
    with get_db() as db:
        db.execute('DELETE FROM users WHERE id=?', (uid,))
    return jsonify({'ok': True})

@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    name = (data.get('name') or '').strip()
    pin  = (data.get('pin') or '').strip()
    with get_db() as db:
        user = db.execute(
            'SELECT id, name, role FROM users WHERE name=? AND pin=?', (name, pin)
        ).fetchone()
    if not user:
        return jsonify({'error': 'Fel namn eller PIN'}), 401
    session['user'] = dict(user)
    return jsonify(dict(user))

@app.route('/api/logout', methods=['POST'])
def logout():
    session.pop('user', None)
    return jsonify({'ok': True})

@app.route('/api/me')
def me():
    return jsonify(session.get('user'))

# ── Machines ─────────────────────────────────────────────────

@app.route('/api/machines')
def list_machines():
    with get_db() as db:
        rows = db.execute('SELECT * FROM machines ORDER BY nr').fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/machines/upload', methods=['POST'])
def upload_machines():
    if 'file' not in request.files:
        return jsonify({'error': 'Ingen fil'}), 400
    f = request.files['file']
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Find header row – stöder både Servicelista ("intern ref") och generiskt format ("maskinnummer")
    header_row = None
    servicelista_format = False
    for i, row in enumerate(rows):
        cells = [str(c or '').strip().lower() for c in row]
        if any('intern ref' in c for c in cells):
            header_row = i
            servicelista_format = True
            break
        if any('maskinnummer' in c for c in cells):
            header_row = i
            break
    if header_row is None:
        return jsonify({'error': 'Kolumn för maskinnummer saknas. Filen ska ha "Intern ref." eller "Maskinnummer".'}), 400

    hdrs_raw = [str(c or '').strip() for c in rows[header_row]]
    hdrs     = [h.lower() for h in hdrs_raw]
    def col(name): return next((i for i,h in enumerate(hdrs) if name in h), None)

    def cell(row, idx):
        if idx is None or idx >= len(row): return ''
        v = row[idx]
        if v is None: return ''
        s = str(v).strip()
        # Rensa float-format på siffror (t.ex. "106590.0" → "106590")
        if s.endswith('.0'):
            try: s = str(int(float(s)))
            except: pass
        return s

    if servicelista_format:
        cMask = col('intern ref')
        cKund = col('kundnamn')
        cAdr  = col('hämtadress rad 1')
        cSta  = col('hämtadress rad 2')
        cKont = col('kontaktperson vid problem')
        cTel  = col('telefonnummer')
        cFab  = col('märke')
        cMod  = col('typ')
        cTill = col('tillverkningsnummer')
        cArs  = col('årsmodell')
        cBesk = col('beskrivning')  # Kolumn "Beskrivning"
        cAgar = col('företag')  # Kolumn A – ägare/fakturakund

        machines = []
        for row in rows[header_row + 1:]:
            nr = cell(row, cMask)
            if not nr or nr.lower() == 'nan': continue
            adress = cell(row, cAdr)
            stad   = cell(row, cSta)
            machines.append({
                'nr':          nr,
                'kund':        cell(row, cKund),
                'anlaggning':  f"{adress}, {stad}".strip(', ') if stad else adress,
                'fabrikat':    cell(row, cFab),
                'modell':      cell(row, cMod),
                'adress':      adress,
                'stad':        stad,
                'kontakt':     cell(row, cKont),
                'telefon':     cell(row, cTel),
                'tillvnr':     cell(row, cTill),
                'arsmodell':   cell(row, cArs),
                'beskrivning': cell(row, cBesk),
                'agare':       cell(row, cAgar),
                'inkopar':     '',
                'notering':    '',
            })
    else:
        cMask = col('maskinnummer')
        cKund = col('kund')
        cAnl  = col('anläggning') or col('anlaggning')
        cFab  = col('fabrikat')
        cMod  = col('modell')
        cInk  = col('inköp') or col('inkop')
        cNot  = col('anteckn')
        cBesk = col('beskrivning')

        machines = []
        for row in rows[header_row + 1:]:
            nr = cell(row, cMask)
            if not nr: continue
            machines.append({
                'nr':          nr,
                'kund':        cell(row, cKund),
                'anlaggning':  cell(row, cAnl),
                'fabrikat':    cell(row, cFab),
                'modell':      cell(row, cMod),
                'inkopar':     cell(row, cInk),
                'notering':    cell(row, cNot),
                'beskrivning': cell(row, cBesk),
                'adress': '', 'stad': '', 'kontakt': '', 'telefon': '', 'tillvnr': '', 'arsmodell': '', 'agare': '',
            })

    with get_db() as db:
        db.executemany("""
            INSERT INTO machines
                (nr, kund, anlaggning, fabrikat, modell, inkopar, notering, beskrivning,
                 adress, stad, kontakt, telefon, tillvnr, arsmodell, agare)
            VALUES
                (:nr, :kund, :anlaggning, :fabrikat, :modell, :inkopar, :notering, :beskrivning,
                 :adress, :stad, :kontakt, :telefon, :tillvnr, :arsmodell, :agare)
            ON CONFLICT(nr) DO UPDATE SET
                kund=excluded.kund, anlaggning=excluded.anlaggning,
                fabrikat=excluded.fabrikat, modell=excluded.modell,
                inkopar=excluded.inkopar, notering=excluded.notering,
                beskrivning=excluded.beskrivning,
                adress=excluded.adress, stad=excluded.stad,
                kontakt=excluded.kontakt, telefon=excluded.telefon,
                tillvnr=excluded.tillvnr, arsmodell=excluded.arsmodell,
                agare=excluded.agare
        """, machines)

    return jsonify({'imported': len(machines)})

@app.route('/api/machines/<nr>', methods=['GET'])
def get_machine(nr):
    with get_db() as db:
        row = db.execute('SELECT * FROM machines WHERE UPPER(nr)=UPPER(?)', (nr,)).fetchone()
    if not row:
        return jsonify(None), 404
    return jsonify(dict(row))

# ── Leveransgodkännande ──────────────────────────────────────

def init_leverans_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS leverans (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            saved_at  TEXT,
            tekniker  TEXT,
            data_json TEXT
        )''')

init_leverans_db()

@app.route('/api/leverans', methods=['POST'])
def save_leverans():
    data = request.get_json()
    with get_db() as db:
        db.execute('INSERT INTO leverans (saved_at, tekniker, data_json) VALUES (?,?,?)',
                   (datetime.now().isoformat(), data.get('tekniker',''), json.dumps(data, ensure_ascii=False)))
    return jsonify({'ok': True})

@app.route('/api/leverans', methods=['GET'])
def list_leverans():
    with get_db() as db:
        rows = db.execute('SELECT id, saved_at, tekniker, data_json FROM leverans ORDER BY saved_at DESC').fetchall()
    return jsonify([{**dict(r), 'data': json.loads(r['data_json'])} for r in rows])

# ── Protocols ────────────────────────────────────────────────

@app.route('/api/protocols/search')
def search_protocols():
    limit  = min(int(request.args.get('limit', 100)), 500)
    kund   = request.args.get('kund', '').strip()
    from_d = request.args.get('from', '').strip()
    to_d   = request.args.get('to', '').strip()

    query  = 'SELECT * FROM protocols WHERE 1=1'
    params = []
    if kund:
        query += ' AND LOWER(kund) = LOWER(?)'
        params.append(kund)
    if from_d:
        query += ' AND datum >= ?'
        params.append(from_d)
    if to_d:
        query += ' AND datum <= ?'
        params.append(to_d)
    query += ' ORDER BY datum DESC, saved_at DESC LIMIT ?'
    params.append(limit)

    with get_db() as db:
        rows = db.execute(query, params).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d['items']          = json.loads(d['items_json'] or '[]')
        d['resolved']       = json.loads(d.get('resolved_json') or '{}')
        d['workflow_log']   = json.loads(d.get('workflow_log') or '[]')
        if not d.get('workflow_status'): d['workflow_status'] = 'inkommen'
        del d['items_json'], d['resolved_json']
        result.append(d)
    return jsonify(result)

@app.route('/api/protocols/recent')
def recent_protocols():
    limit = min(int(request.args.get('limit', 15)), 50)
    with get_db() as db:
        rows = db.execute(
            'SELECT * FROM protocols ORDER BY saved_at DESC LIMIT ?', (limit,)
        ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d['items']    = json.loads(d['items_json'] or '[]')
        d['resolved']      = json.loads(d.get('resolved_json') or '{}')
        d['workflow_log']  = json.loads(d.get('workflow_log') or '[]')
        if 'workflow_status' not in d or not d['workflow_status']:
            d['workflow_status'] = 'inkommen'
        del d['items_json'], d['resolved_json']
        result.append(d)
    return jsonify(result)

@app.route('/api/protocols/<maskin_nr>')
def get_protocols(maskin_nr):
    with get_db() as db:
        rows = db.execute(
            'SELECT * FROM protocols WHERE UPPER(maskin_nr)=UPPER(?) ORDER BY saved_at DESC LIMIT 30',
            (maskin_nr,)
        ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d['items']          = json.loads(d['items_json'] or '[]')
        d['resolved']       = json.loads(d.get('resolved_json') or '{}')
        d['workflow_log']   = json.loads(d.get('workflow_log') or '[]')
        if not d.get('workflow_status'):
            d['workflow_status'] = 'inkommen'
        del d['items_json'], d['resolved_json']
        result.append(d)
    return jsonify(result)

@app.route('/api/protocols/<int:protocol_id>', methods=['DELETE'])
def delete_protocol(protocol_id):
    with get_db() as db:
        db.execute('DELETE FROM protocols WHERE id=?', (protocol_id,))
    return jsonify({'ok': True})

@app.route('/api/leverans/<int:leverans_id>', methods=['DELETE'])
def delete_leverans(leverans_id):
    with get_db() as db:
        db.execute('DELETE FROM leverans WHERE id=?', (leverans_id,))
    return jsonify({'ok': True})

@app.route('/api/protocols/<int:protocol_id>/resolve', methods=['PATCH'])
def resolve_item(protocol_id):
    data = request.get_json()
    nr   = data.get('nr')
    if not nr:
        return jsonify({'error': 'nr krävs'}), 400
    with get_db() as db:
        row = db.execute('SELECT resolved_json FROM protocols WHERE id=?', (protocol_id,)).fetchone()
        if not row:
            return jsonify({'error': 'Protokoll ej hittat'}), 404
        resolved = json.loads(row['resolved_json'] or '{}')
        if data.get('undo'):
            resolved.pop(nr, None)
        else:
            resolved[nr] = {
                'date': datetime.now().strftime('%Y-%m-%d'),
                'by':   data.get('by', '').strip()
            }
        db.execute('UPDATE protocols SET resolved_json=? WHERE id=?',
                   (json.dumps(resolved, ensure_ascii=False), protocol_id))
    return jsonify({'resolved': resolved})

# ── Workflow ──────────────────────────────────────────────────

WORKFLOW_STEPS = ['inkommen', 'hanterad']
WORKFLOW_LABELS = {
    'inkommen': 'Inkommen',
    'hanterad': 'Hanterad',
}

@app.route('/api/protocols/<int:protocol_id>/workflow', methods=['PATCH'])
def update_workflow(protocol_id):
    data   = request.get_json()
    status = data.get('status')
    note   = data.get('note', '').strip()
    user   = session.get('user', {})
    if status not in WORKFLOW_STEPS:
        return jsonify({'error': 'Ogiltigt status'}), 400
    with get_db() as db:
        row = db.execute('SELECT workflow_log FROM protocols WHERE id=?', (protocol_id,)).fetchone()
        if not row:
            return jsonify({'error': 'Protokoll ej hittat'}), 404
        log = json.loads(row['workflow_log'] or '[]')
        log.append({
            'status': status,
            'label':  WORKFLOW_LABELS[status],
            'by':     user.get('name', 'Okänd'),
            'at':     datetime.now().strftime('%Y-%m-%d %H:%M'),
            'note':   note,
        })
        db.execute('UPDATE protocols SET workflow_status=?, workflow_log=? WHERE id=?',
                   (status, json.dumps(log, ensure_ascii=False), protocol_id))
    return jsonify({'workflow_status': status, 'workflow_log': log})

@app.route('/api/protocols', methods=['POST'])
def save_protocol():
    data = request.get_json()
    if not data or not data.get('maskin_nr'):
        return jsonify({'error': 'maskin_nr krävs'}), 400
    with get_db() as db:
        db.execute("""
            INSERT INTO protocols (maskin_nr, datum, type, machine_type, kund, anlaggning, tekniker, modell, items_json, saved_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (
            data['maskin_nr'].upper(),
            data.get('datum', ''),
            data.get('type', ''),
            data.get('machine_type', ''),
            data.get('kund', ''),
            data.get('anlaggning', ''),
            data.get('tekniker', ''),
            data.get('modell', ''),
            json.dumps(data.get('items', []), ensure_ascii=False),
            datetime.now().isoformat()
        ))
        pid = db.execute('SELECT last_insert_rowid()').fetchone()[0]
    return jsonify({'id': pid}), 201

# ── Photo upload ─────────────────────────────────────────────

ALLOWED_EXT = {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.heic'}

@app.route('/api/photos', methods=['POST'])
def upload_photo():
    if 'photo' not in request.files:
        return jsonify({'error': 'Ingen fil'}), 400
    f = request.files['photo']
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({'error': 'Otillåtet filformat'}), 400
    filename = uuid.uuid4().hex + ext
    f.save(os.path.join(UPLOADS, filename))
    return jsonify({'filename': filename}), 201

@app.route('/uploads/<filename>')
def serve_upload(filename):
    return send_from_directory(UPLOADS, filename)

if __name__ == '__main__':
    import socket
    host = '0.0.0.0'
    port = 5000
    local_ip = socket.gethostbyname(socket.gethostname())
    print(f'\n  Serviceprotokoll-server startad!')
    print(f'  Öppna i webbläsare: http://localhost:{port}')
    print(f'  Från andra enheter: http://{local_ip}:{port}')
    print(f'  Tryck Ctrl+C för att stänga av\n')
    app.run(host=host, port=port, debug=False)
