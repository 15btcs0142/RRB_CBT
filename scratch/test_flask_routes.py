import os
import sys
import types
import sqlite3

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
os.chdir(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Mock pdfkit if not installed
try:
    import pdfkit
except ImportError:
    sys.modules['pdfkit'] = types.ModuleType('pdfkit')

import app

print("Initializing test client...")
app.app.config['TESTING'] = True
app.app.config['WTF_CSRF_ENABLED'] = False
client = app.app.test_client()

print("\n--- Testing Public and Student Routes ---")
routes_to_test = [
    ('/', 'GET'),
    ('/waiting', 'GET'),
    ('/exam', 'GET'),
    ('/submitted', 'GET'),
    ('/admin/login', 'GET'),
    ('/teacher/login', 'GET'),
    ('/admin/dashboard', 'GET'),
    ('/teacher/dashboard', 'GET'),
    ('/teacher/create-test', 'GET'),
    ('/admin/create-test', 'GET'),
    ('/api/test_numbers', 'GET'),
    ('/api/check_test_no', 'GET'),
    ('/api/exam_status', 'GET'),
    ('/admin/monitoring', 'GET'),
    ('/admin/questions', 'GET'),
    ('/admin/students', 'GET'),
    ('/admin/results', 'GET'),
    ('/admin/settings', 'GET'),
    ('/admin/bulletins', 'GET'),
    ('/admin/teachers', 'GET'),
    ('/admin/scheduled_tests', 'GET'),
    ('/teacher/students', 'GET'),
    ('/teacher/monitoring', 'GET'),
    ('/teacher/history', 'GET'),
    ('/teacher/profile', 'GET'),
]

for url, method in routes_to_test:
    try:
        if method == 'GET':
            res = client.get(url)
        else:
            res = client.post(url)
        print(f"[{method}] {url:<30} -> Status: {res.status_code}")
    except Exception as e:
        print(f"[CRASH BUG] [{method}] {url:<30} -> EXCEPTION: {e}")

print("\n--- Testing Teacher Authenticated Routes ---")
with client.session_transaction() as sess:
    sess['teacher_logged_in'] = True
    sess['teacher_id'] = 1
    sess['teacher_name'] = 'Test Teacher'

teacher_routes = [
    ('/teacher/dashboard', 'GET'),
    ('/teacher/create-test', 'GET'),
    ('/teacher/students', 'GET'),
    ('/teacher/monitoring', 'GET'),
    ('/teacher/history', 'GET'),
    ('/teacher/profile', 'GET'),
    ('/teacher/print_test/1', 'GET'),
]

for url, method in teacher_routes:
    try:
        res = client.get(url)
        print(f"[TEACHER AUTH] {url:<30} -> Status: {res.status_code}")
    except Exception as e:
        print(f"[CRASH BUG] [TEACHER AUTH] {url:<30} -> EXCEPTION: {e}")

print("\n--- Testing Admin Authenticated Routes ---")
with client.session_transaction() as sess:
    sess['admin_logged_in'] = True
    sess['admin_username'] = 'admin'

admin_routes = [
    ('/admin/dashboard', 'GET'),
    ('/admin/create-test', 'GET'),
    ('/admin/questions', 'GET'),
    ('/admin/students', 'GET'),
    ('/admin/results', 'GET'),
    ('/admin/settings', 'GET'),
    ('/admin/bulletins', 'GET'),
    ('/admin/teachers', 'GET'),
    ('/admin/scheduled_tests', 'GET'),
]

for url, method in admin_routes:
    try:
        res = client.get(url)
        print(f"[ADMIN AUTH] {url:<30} -> Status: {res.status_code}")
    except Exception as e:
        print(f"[CRASH BUG] [ADMIN AUTH] {url:<30} -> EXCEPTION: {e}")
