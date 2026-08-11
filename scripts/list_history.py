import sqlite3, json, os

DB='database.db'
if not os.path.exists(DB):
    print('ERROR: database.db not found in', os.getcwd())
    raise SystemExit(1)

conn=sqlite3.connect(DB)
conn.row_factory=sqlite3.Row
c=conn.cursor()
c.execute('SELECT id, teacher_id, class, section, subject, chapter, test_no, created_at FROM test_generation_history ORDER BY id')
rows=c.fetchall()
print(json.dumps([dict(r) for r in rows], default=str, indent=2))
conn.close()
