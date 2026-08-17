import os
import sys
import json
import unittest

sys.path.insert(0, r'c:\Users\atuls\OneDrive\Desktop\RRB_CBT\RRB_cbt-v1.10\RRB_v110')
from app import app, init_db, generate_jwt_token, decode_jwt_token

class TestJWTAuthentication(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        app.config['SECRET_KEY'] = 'test_secret_key'
        self.client = app.test_client()
        with app.app_context():
            init_db()

    def test_jwt_token_generation_and_decoding(self):
        """Test low-level token generation and PyJWT decoding."""
        token, exp = generate_jwt_token(identity="admin_test", role="admin", name="Test Admin")
        self.assertIsNotNone(token)
        
        payload = decode_jwt_token(token)
        self.assertEqual(payload['sub'], "admin_test")
        self.assertEqual(payload['role'], "admin")
        self.assertEqual(payload['name'], "Test Admin")

    def test_admin_token_issuance_route(self):
        """Test POST /api/token for admin role."""
        res = self.client.post('/api/token', json={
            'role': 'admin',
            'username': 'admin',
            'password': 'admin123'
        })
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data['status'], 'success')
        self.assertEqual(data['token_type'], 'Bearer')
        self.assertIn('access_token', data)

    def test_invalid_admin_password(self):
        """Test POST /api/token with wrong password."""
        res = self.client.post('/api/token', json={
            'role': 'admin',
            'username': 'admin',
            'password': 'wrongpassword'
        })
        self.assertEqual(res.status_code, 401)
        data = res.get_json()
        self.assertEqual(data['status'], 'error')

    def test_student_token_issuance_route(self):
        """Test POST /api/token for student role."""
        res = self.client.post('/api/token', json={
            'role': 'student',
            'student_id': 'STU_1001',
            'name': 'Gaurav Student',
            'class': '10',
            'section': 'A'
        })
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data['status'], 'success')
        self.assertEqual(data['user']['id'], 'STU_1001')

    def test_jwt_required_decorator_protection(self):
        """Test /api/user/me protection with @jwt_required."""
        # 1. Without header -> 401
        res_no_auth = self.client.get('/api/user/me')
        self.assertEqual(res_no_auth.status_code, 401)

        # 2. Invalid token -> 401
        res_bad_auth = self.client.get('/api/user/me', headers={'Authorization': 'Bearer invalid_token_xyz'})
        self.assertEqual(res_bad_auth.status_code, 401)

        # 3. Valid token -> 200
        token, _ = generate_jwt_token(identity="admin", role="admin", name="System Admin")
        res_valid = self.client.get('/api/user/me', headers={'Authorization': f'Bearer {token}'})
        self.assertEqual(res_valid.status_code, 200)
        data = res_valid.get_json()
        self.assertEqual(data['status'], 'success')
        self.assertEqual(data['user']['id'], 'admin')

    def test_web_routes_unaffected(self):
        """Verify web routes continue to load cleanly."""
        res = self.client.get('/admin')
        self.assertIn(res.status_code, [200, 302])
        
        res_teacher = self.client.get('/teacher/login')
        self.assertIn(res_teacher.status_code, [200, 302])

if __name__ == '__main__':
    unittest.main()
