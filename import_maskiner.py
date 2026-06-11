import pandas as pd, sqlite3

path_xl = r'C:\Users\FelixLiljegren\Downloads\Servicelista Liljegren Rätt.xlsx'
path_db = r'C:\Users\FelixLiljegren\OneDrive - Liljegrens i Tyresö AB\Claude\Serviceprotokoll\serviceprotokoll.db'

df = pd.read_excel(path_xl, header=0, dtype=str)
df = df.fillna('')

conn = sqlite3.connect(path_db)
conn.execute('''CREATE TABLE IF NOT EXISTS machines (
    nr TEXT PRIMARY KEY, kund TEXT DEFAULT '', anlaggning TEXT DEFAULT '',
    fabrikat TEXT DEFAULT '', modell TEXT DEFAULT '',
    inkopar TEXT DEFAULT '', notering TEXT DEFAULT '')''')

for col, typ in [('adress','TEXT'),('stad','TEXT'),('kontakt','TEXT'),
                 ('telefon','TEXT'),('tillvnr','TEXT'),('arsmodell','TEXT'),('beskrivning','TEXT')]:
    try:
        conn.execute(f'ALTER TABLE machines ADD COLUMN {col} {typ} DEFAULT ""')
    except Exception:
        pass
conn.commit()

imported = 0
for _, row in df.iterrows():
    raw_nr = row['Intern ref.'].strip()
    try: nr = str(int(float(raw_nr)))
    except: nr = raw_nr
    if not nr or nr == 'nan':
        continue
    kund    = row['Placeringsinfo: Kundnamn'].strip()
    adress  = row['Placeringsinfo: Hämtadress rad 1'].strip()
    stad    = row['Placeringsinfo: Hämtadress rad 2'].strip()
    kontakt = row['Placeringsinfo: Kontaktperson vid problem'].strip()
    telefon = row['Placeringsinfo: Telefonnummer - Kontaktperson'].strip()
    marke   = row['Utrustningsinfo: Märke'].strip()
    typ     = row['Utrustningsinfo: Typ'].strip()
    try: tillvnr = str(int(float(row['Utrustningsinfo: Tillverkningsnummer'].strip()))) if row['Utrustningsinfo: Tillverkningsnummer'].strip() else ''
    except: tillvnr = row['Utrustningsinfo: Tillverkningsnummer'].strip()
    try: arsmod = str(int(float(row['Utrustningsinfo: Årsmodell'].strip()))) if row['Utrustningsinfo: Årsmodell'].strip() else ''
    except: arsmod = row['Utrustningsinfo: Årsmodell'].strip()
    # ↓ Justera kolumnnamnet nedan om det skiljer sig i din Excel-fil
    BESKRIVNING_KOLUMN = 'Utrustningsinfo: Beskrivning'
    beskr   = row[BESKRIVNING_KOLUMN].strip() if BESKRIVNING_KOLUMN in row else ''
    anl     = f'{adress}, {stad}' if stad else adress

    conn.execute('''INSERT INTO machines
        (nr,kund,anlaggning,fabrikat,modell,adress,stad,kontakt,telefon,tillvnr,arsmodell,beskrivning)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(nr) DO UPDATE SET
            kund=excluded.kund, anlaggning=excluded.anlaggning,
            fabrikat=excluded.fabrikat, modell=excluded.modell,
            adress=excluded.adress, stad=excluded.stad,
            kontakt=excluded.kontakt, telefon=excluded.telefon,
            tillvnr=excluded.tillvnr, arsmodell=excluded.arsmodell,
            beskrivning=excluded.beskrivning''',
        (nr, kund, anl, marke, typ, adress, stad, kontakt, telefon, tillvnr, arsmod, beskr))
    imported += 1

conn.commit()
conn.close()
print(f'Importerade {imported} maskiner OK')
