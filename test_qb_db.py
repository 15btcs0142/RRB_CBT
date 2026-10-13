import sys
sys.path.insert(0, '.')
import app
app.init_db()
import sqlite3
conn = app.get_db()
c = conn.cursor()
c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'question_bank%'")
tables = [r[0] for r in c.fetchall()]
conn.close()
print('[OK] Question Bank tables:', tables)
assert 'question_bank_papers' in tables
assert 'question_bank_items' in tables
print('[ALL OK] DB schema verified successfully')
