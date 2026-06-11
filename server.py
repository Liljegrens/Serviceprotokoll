from flask import Flask, jsonify, request, send_from_directory, session
import sqlite3, json, os, secrets
from datetime import datetime
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
DB   = os.path.join(BASE, 'serviceprotokoll.db')
app  = Flask(__name__, static_folder=BASE)
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(32))

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
            avdelning  TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS protocols (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            maskin_nr     TEXT NOT NULL,
            datum         TEXT,
            type          TEXT,
            kund          TEXT,
            anlaggning    TEXT,
            tekniker      TEXT,
            modell        TEXT,
            items_json    TEXT,
            saved_at      TEXT,
            resolved_json TEXT DEFAULT '{}'
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

# ── Serve frontend ───────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(BASE, 'serviceprotokoll.html')

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

    # Find header row
    header_row = None
    for i, row in enumerate(rows):
        cells = [str(c or '').strip().lower() for c in row]
        if any('maskinnummer' in c for c in cells):
            header_row = i
            break
    if header_row is None:
        return jsonify({'error': 'Kolumn "Maskinnummer" saknas'}), 400

    hdrs = [str(c or '').strip().lower() for c in rows[header_row]]
    def col(name): return next((i for i,h in enumerate(hdrs) if name in h), None)

    cMask = col('maskinnummer')
    cKund = col('kund')
    cAnl  = col('anläggning') or col('anlaggning')
    cFab  = col('fabrikat')
    cMod  = col('modell')
    cInk  = col('inköp') or col('inkop')
    cNot  = col('anteckn')

    machines = []
    for row in rows[header_row + 1:]:
        nr = str(row[cMask] or '').strip() if cMask is not None else ''
        if not nr:
            continue
        machines.append({
            'nr':         nr,
            'kund':       str(row[cKund] or '').strip() if cKund is not None else '',
            'anlaggning': str(row[cAnl]  or '').strip() if cAnl  is not None else '',
            'fabrikat':   str(row[cFab]  or '').strip() if cFab  is not None else '',
            'modell':     str(row[cMod]  or '').strip() if cMod  is not None else '',
            'inkopar':    str(row[cInk]  or '').strip() if cInk  is not None else '',
            'notering':   str(row[cNot]  or '').strip() if cNot  is not None else '',
        })

    with get_db() as db:
        db.executemany("""
            INSERT INTO machines (nr, kund, anlaggning, fabrikat, modell, inkopar, notering)
            VALUES (:nr, :kund, :anlaggning, :fabrikat, :modell, :inkopar, :notering)
            ON CONFLICT(nr) DO UPDATE SET
                kund=excluded.kund, anlaggning=excluded.anlaggning,
                fabrikat=excluded.fabrikat, modell=excluded.modell,
                inkopar=excluded.inkopar, notering=excluded.notering
        """, machines)

    return jsonify({'imported': len(machines)})

@app.route('/api/machines/<nr>', methods=['GET'])
def get_machine(nr):
    with get_db() as db:
        row = db.execute('SELECT * FROM machines WHERE UPPER(nr)=UPPER(?)', (nr,)).fetchone()
    if not row:
        return jsonify(None), 404
    return jsonify(dict(row))

# ── Protocols ────────────────────────────────────────────────

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
        d['items'] = json.loads(d['items_json'] or '[]')
        d['resolved'] = json.loads(d.get('resolved_json') or '{}')
        del d['items_json']
        del d['resolved_json']
        result.append(d)
    return jsonify(result)

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

@app.route('/api/protocols', methods=['POST'])
def save_protocol():
    data = request.get_json()
    if not data or not data.get('maskin_nr'):
        return jsonify({'error': 'maskin_nr krävs'}), 400
    with get_db() as db:
        db.execute("""
            INSERT INTO protocols (maskin_nr, datum, type, kund, anlaggning, tekniker, modell, items_json, saved_at)
            VALUES (?,?,?,?,?,?,?,?,?)
        """, (
            data['maskin_nr'].upper(),
            data.get('datum', ''),
            data.get('type', ''),
            data.get('kund', ''),
            data.get('anlaggning', ''),
            data.get('tekniker', ''),
            data.get('modell', ''),
            json.dumps(data.get('items', []), ensure_ascii=False),
            datetime.now().isoformat()
        ))
        pid = db.execute('SELECT last_insert_rowid()').fetchone()[0]
    return jsonify({'id': pid}), 201

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
