import os
import sys
import unittest

# Add app directory to sys.path
sys.path.insert(0, r'c:\Users\atuls\OneDrive\Desktop\RRB_CBT\RRB_cbt-v1.10\RRB_v110')
from app import app, init_db, get_db, save_paper_to_repository

class TestRRBUpdates(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        app.config['SECRET_KEY'] = 'test_secret_key'
        self.client = app.test_client()
        with app.app_context():
            init_db()

    def test_database_tables_exist(self):
        """Verify that mcq_test_history and teacher_notifications tables exist."""
        with app.app_context():
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='mcq_test_history'")
            self.assertIsNotNone(c.fetchone())
            
            c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='teacher_notifications'")
            self.assertIsNotNone(c.fetchone())
            conn.close()

    def test_paper_repository_saving(self):
        """Verify paper saving into paper/Class_<N>/Section_<X>/ folder hierarchy."""
        sample_paper = {
            "sections": [
                {
                    "section_label": "Section A",
                    "section_title": "Multiple Choice Questions",
                    "instruction": "Answer all questions",
                    "questions": [
                        {"number": 1, "question": "What is 2+2?", "options": ["(a) 3", "(b) 4", "(c) 5", "(d) 6"]}
                    ]
                }
            ]
        }
        
        abs_file, filename, rel_link = save_paper_to_repository(
            sample_paper, "Class 10", "Section A", "Mathematics", "Unit Test 1", "Gaurav Shukla"
        )
        
        self.assertTrue(os.path.exists(abs_file))
        self.assertIn("paper", abs_file)
        self.assertIn("Class_Class_10", abs_file)
        self.assertIn("Section_Section_A", abs_file)
        self.assertTrue(filename.endswith('.doc'))

if __name__ == '__main__':
    unittest.main()
