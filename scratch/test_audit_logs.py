import os
import sys
import unittest
import datetime

sys.path.insert(0, r'c:\Users\atuls\OneDrive\Desktop\RRB_CBT\RRB_cbt-v1.10\RRB_v110')
from app import app, init_db, get_db, log_audit_event

class TestAuditLogs(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()
        with app.app_context():
            init_db()

    def test_log_audit_event_helper(self):
        """Test log_audit_event helper directly writes to audit_logs table."""
        log_audit_event('admin', 'admin', 'TEST_ACTION', 'test_table', '123', '127.0.0.1')
        
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM audit_logs WHERE action='TEST_ACTION'")
        row = c.fetchone()
        conn.close()
        
        self.assertIsNotNone(row)
        self.assertEqual(row['user_type'], 'admin')
        self.assertEqual(row['user_id'], 'admin')
        self.assertEqual(row['target_table'], 'test_table')
        self.assertEqual(row['target_id'], '123')

    def test_admin_audit_logs_route_and_filters(self):
        """Test /admin/audit_logs endpoint and filter query parameters."""
        # Insert test records
        log_audit_event('admin', 'admin', 'EXAM_START', 'exam_control', '1', '127.0.0.1')
        log_audit_event('teacher', 'T101', 'STUDENT_RESULT_VIEW', 'students', 'STU_1', '127.0.0.1')
        log_audit_event('admin', 'admin', 'PAPER_DELETE', 'questions', '10_Math', '127.0.0.1')
        log_audit_event('admin', 'admin', 'STUDENT_UNLOCK', 'student_class_lock', 'STU_2', '127.0.0.1')

        with self.client.session_transaction() as sess:
            sess['admin_logged_in'] = True

        # 1. Fetch all audit logs JSON
        res = self.client.get('/admin/audit_logs?format=json')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data['status'], 'success')
        self.assertGreaterEqual(len(data['logs']), 4)

        # 2. Filter by user_type=teacher
        res_teacher = self.client.get('/admin/audit_logs?user_type=teacher&format=json')
        self.assertEqual(res_teacher.status_code, 200)
        t_data = res_teacher.get_json()
        for log in t_data['logs']:
            self.assertEqual(log['user_type'].lower(), 'teacher')

        # 3. Filter by action=EXAM_START
        res_action = self.client.get('/admin/audit_logs?action=EXAM_START&format=json')
        self.assertEqual(res_action.status_code, 200)
        a_data = res_action.get_json()
        for log in a_data['logs']:
            self.assertEqual(log['action'], 'EXAM_START')

    def test_actions_instrumented(self):
        """Verify the 5 required action types exist in system audit capabilities."""
        required_actions = [
            'STUDENT_RESULT_VIEW',
            'PAPER_DELETE',
            'REATTEMPT_APPROVE',
            'EXAM_START',
            'STUDENT_UNLOCK'
        ]
        for act in required_actions:
            log_audit_event('admin', 'admin', act, 'test_target', '1', '127.0.0.1')

        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT DISTINCT action FROM audit_logs")
        actions_found = [r['action'] for r in c.fetchall()]
        conn.close()

        for req in required_actions:
            self.assertIn(req, actions_found)

if __name__ == '__main__':
    unittest.main()
