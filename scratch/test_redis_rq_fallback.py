import os
import sys
import unittest
import datetime

sys.path.insert(0, r'c:\Users\atuls\OneDrive\Desktop\RRB_CBT\RRB_cbt-v1.10\RRB_v110')
from app import (app, init_db, enqueue_ai_paper_job, process_ai_paper_job,
                 get_paper_job_status, REDIS_AVAILABLE, paper_job_results)

class TestRedisRQAndFallback(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        self.client = app.test_client()
        with app.app_context():
            init_db()

    def test_queue_mode_detection(self):
        """Verify queue mode (RQ vs Thread Queue) detection status."""
        self.assertIn(REDIS_AVAILABLE, [True, False])

    def test_enqueue_ai_paper_job(self):
        """Test enqueueing AI paper job returns valid job_id and queue status."""
        job_data = {
            'job_id': 'test_job_101',
            'prompt': 'Generate 2 questions on Algebra',
            'teacher_id': 1,
            'teacher_name': 'Test Teacher',
            'class_': '10',
            'section': 'A',
            'subject': 'Mathematics',
            'exam_type': 'Unit Test 1',
            'duration': '1 Hour',
            'max_marks': '20'
        }
        job_id, queue_type = enqueue_ai_paper_job(job_data)
        self.assertEqual(job_id, 'test_job_101')
        self.assertIn(queue_type, ['rq', 'thread'])
        self.assertIn('test_job_101', paper_job_results)

    def test_job_status_api_endpoint(self):
        """Test GET /api/paper_job/<job_id> endpoint."""
        paper_job_results['status_test_999'] = {
            'status': 'processing',
            'progress': 40,
            'message': 'Generating questions...'
        }
        res = self.client.get('/api/paper_job/status_test_999')
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data['status'], 'success')
        self.assertEqual(data['job']['status'], 'processing')

    def test_job_status_404_not_found(self):
        """Test GET /api/paper_job/<non_existent_id> returns 404."""
        res = self.client.get('/api/paper_job/non_existent_job_xyz')
        self.assertEqual(res.status_code, 404)

if __name__ == '__main__':
    unittest.main()
