import os
import sys
import types
import sqlite3

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
os.chdir(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import app
from validators import validate_field, STUDENT_LOGIN_SCHEMA

print("==================================================")
print("       RUNNING BUG FIX VERIFICATION SUITE         ")
print("==================================================")

# 1. Test Student Name validation
print("\n--- 1. Testing Student Name Validation ---")
test_names = ["O'Connor", "D'Souza", "John Doe-Smith", "Rahul Sharma", "अतुल कुमार"]
for name in test_names:
    try:
        val = validate_field(name, 'name', STUDENT_LOGIN_SCHEMA['name'])
        print(f"  [PASS] Name '{name}' successfully validated -> '{val}'")
    except Exception as e:
        print(f"  [FAIL] Name '{name}' failed validation: {e}")

# 2. Test Admin Login with /admin/login
print("\n--- 2. Testing Admin Login Route ---")
app.app.config['TESTING'] = True
client = app.app.test_client()

res = client.get('/admin/login')
print(f"  GET /admin/login status code: {res.status_code} (Expected 200)")

res_post = client.post('/admin/login', data={'username': 'admin', 'password': 'admin123'})
print(f"  POST /admin/login status code: {res_post.status_code} (Expected 302 redirect to /admin/dashboard)")

# 3. Test Teacher Create Test GET route (No NameError)
print("\n--- 3. Testing Teacher Create Test GET Route ---")
with client.session_transaction() as sess:
    sess['teacher_logged_in'] = True
    sess['teacher_id'] = 1
    sess['teacher_name'] = 'Test Teacher'

res_teacher = client.get('/teacher/create_test_v2')
print(f"  GET /teacher/create_test_v2 status code: {res_teacher.status_code} (Expected 200)")

# 4. Test Teacher Access Function
print("\n--- 4. Testing teacher_has_access Helper ---")
try:
    has_access = app.teacher_has_access(1, '10th', 'Science')
    print(f"  teacher_has_access(1, '10th', 'Science') returned: {has_access} (No NameError!)")
except Exception as e:
    print(f"  [FAIL] teacher_has_access raised exception: {e}")

# 5. Test Reattempt Scoped Deletion
print("\n--- 5. Testing Multi-Subject Reattempt Isolation ---")
conn = app.get_db()
c = conn.cursor()
c.execute("INSERT OR REPLACE INTO students (student_id, name, class, subject, status) VALUES ('TEST_STU_1', 'Multi Test Student', '10th', 'Science', 'Submitted')")
c.execute("INSERT OR REPLACE INTO questions (id, class, subject, question, option_a, option_b, option_c, option_d, correct_answer) VALUES (9001, '10th', 'Science', 'Q1 Sci', 'A', 'B', 'C', 'D', 'A')")
c.execute("INSERT OR REPLACE INTO questions (id, class, subject, question, option_a, option_b, option_c, option_d, correct_answer) VALUES (9002, '10th', 'Maths', 'Q1 Math', 'A', 'B', 'C', 'D', 'B')")
c.execute("INSERT OR REPLACE INTO responses (student_id, question_id, selected_option) VALUES ('TEST_STU_1', 9001, 'A')")
c.execute("INSERT OR REPLACE INTO responses (student_id, question_id, selected_option) VALUES ('TEST_STU_1', 9002, 'B')")
c.execute("INSERT INTO results (student_id, name, class, subject, score, total_questions, percentage) VALUES ('TEST_STU_1', 'Multi Test Student', '10th', 'Science', 1, 1, 100)")
c.execute("INSERT INTO results (student_id, name, class, subject, score, total_questions, percentage) VALUES ('TEST_STU_1', 'Multi Test Student', '10th', 'Maths', 1, 1, 100)")
c.execute("INSERT INTO reattempt_requests (student_id, class, subject, status) VALUES ('TEST_STU_1', '10th', 'Science', 'pending')")
req_id = c.lastrowid
conn.commit()

# Call admin allow reattempt via test client
with client.session_transaction() as sess:
    sess['admin_logged_in'] = True
    sess['admin_username'] = 'admin'

res_approve = client.post(f'/admin/reattempt_request/{req_id}/approve', json={'note': 'Approved test'})
print(f"  Approve Science reattempt status code: {res_approve.status_code}")

# Verify results: Science result should be gone, Maths result should remain!
c.execute("SELECT subject FROM results WHERE student_id='TEST_STU_1'")
remaining_results = [r['subject'] for r in c.fetchall()]
print(f"  Remaining results for student: {remaining_results}")
if 'Maths' in remaining_results and 'Science' not in remaining_results:
    print("  [PASS] Reattempt scoped deletion successfully preserved Maths results while clearing Science!")
else:
    print(f"  [FAIL] Unexpected results state: {remaining_results}")

# Clean up test records
c.execute("DELETE FROM students WHERE student_id='TEST_STU_1'")
c.execute("DELETE FROM responses WHERE student_id='TEST_STU_1'")
c.execute("DELETE FROM results WHERE student_id='TEST_STU_1'")
c.execute("DELETE FROM questions WHERE id IN (9001, 9002)")
c.execute("DELETE FROM reattempt_requests WHERE id=?", (req_id,))
conn.commit()
conn.close()

print("\n==================================================")
print("          ALL VERIFICATIONS COMPLETED!            ")
print("==================================================")
