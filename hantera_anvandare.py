"""
Kör detta skript för att lägga till / lista / ta bort användare.
  python hantera_anvandare.py
"""
import sqlite3, os

DB = os.path.join(os.path.dirname(__file__), 'serviceprotokoll.db')

def get_db():
    db = sqlite3.connect(DB)
    db.row_factory = sqlite3.Row
    db.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, pin TEXT NOT NULL, role TEXT DEFAULT 'tekniker')""")
    db.commit()
    return db

def list_users(db):
    rows = db.execute('SELECT id, name, pin, role FROM users ORDER BY name').fetchall()
    if not rows:
        print('  (inga användare)')
    for r in rows:
        print(f"  [{r['id']}] {r['name']}  PIN: {r['pin']}  Roll: {r['role']}")

def main():
    db = get_db()
    while True:
        print('\n── Användarhantering ──────────────────')
        list_users(db)
        print('\n  1. Lägg till användare')
        print('  2. Ta bort användare')
        print('  3. Avsluta')
        val = input('\nVal: ').strip()

        if val == '1':
            name = input('Namn: ').strip()
            pin  = input('PIN (minst 4 siffror): ').strip()
            role = input('Roll (tekniker/admin) [tekniker]: ').strip() or 'tekniker'
            if not name or not pin:
                print('Namn och PIN krävs.')
                continue
            db.execute('INSERT INTO users (name, pin, role) VALUES (?,?,?)', (name, pin, role))
            db.commit()
            print(f'✓ {name} tillagd.')

        elif val == '2':
            uid = input('ID att ta bort: ').strip()
            db.execute('DELETE FROM users WHERE id=?', (uid,))
            db.commit()
            print('✓ Borttagen.')

        elif val == '3':
            break

if __name__ == '__main__':
    main()
