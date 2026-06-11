import sqlite3
db = sqlite3.connect('serviceprotokoll.db')
try:
    db.execute("ALTER TABLE protocols ADD COLUMN workflow_status TEXT DEFAULT 'inkommen'")
    print("workflow_status tillagd")
except Exception as e:
    print("workflow_status:", e)
try:
    db.execute("ALTER TABLE protocols ADD COLUMN workflow_log TEXT DEFAULT '[]'")
    print("workflow_log tillagd")
except Exception as e:
    print("workflow_log:", e)
db.commit()
print("Klart")
