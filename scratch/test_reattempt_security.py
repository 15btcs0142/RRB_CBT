import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import app, get_db

with app.test_client() as client:
    # Test checking reattempt status endpoint for student
    with client.session_transaction() as sess:
        sess['student_id'] = 'TEST_STUDENT_001'
        sess['class'] = '10'
        sess['subject'] = 'Mathematics'

    res = client.get('/check_reattempt_status')
    print("Check reattempt status code:", res.status_code)
    print("Response JSON:", res.get_json())
