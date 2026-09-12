import sqlite3, os, re, jinja2

# 1. Test database schema creation and table columns
conn = sqlite3.connect(':memory:')
conn.row_factory = sqlite3.Row

import app
app.init_db(conn=conn) if 'conn' in app.init_db.__code__.co_varnames else None
