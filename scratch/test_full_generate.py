import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import app
import json

with app.app.test_request_context('/descriptive_paper', method='POST', data={
    'board': 'ICSE',
    'class': '10',
    'subject': 'Computer Application',
    'topics': '4. Constructors',
    'full_syllabus': 'no',
    'exam_type': 'Unit Test',
    'duration': '1 Hour',
    'max_marks': '25',
    'mcq_count': '5',
    'mcq_marks': '1',
    'fib_count': '2',
    'fib_marks': '1',
    'tf_count': '0',
    'tf_marks': '1',
    'ar_count': '0',
    'ar_marks': '1',
    'vs_count': '2',
    'vs_marks': '2',
    'sh_count': '2',
    'sh_marks': '3',
    'lg_count': '1',
    'lg_marks': '5',
    'cs_count': '0',
    'cs_marks': '4'
}):
    from flask import session
    session['admin_logged_in'] = True
    resp = app.descriptive_paper()
    raw = resp.get_data(as_text=True)
    data = json.loads(raw)
    print("STATUS:", data.get('status'))
    if data.get('status') == 'success':
        sections = data.get('paper', {}).get('sections', [])
        print(f"Total Sections Generated: {len(sections)}")
        total_questions = 0
        total_marks = 0
        for s in sections:
            q_list = s.get('questions', [])
            m_each = s.get('marks_per_question', 1)
            total_questions += len(q_list)
            total_marks += len(q_list) * m_each
            print(f"  • {s.get('section_label')}: {s.get('section_title')} -> {len(q_list)} Questions ({m_each}M each)")
            for q in q_list:
                print(f"     [Q{q.get('number')}] {q.get('question')[:75]}...")
        print(f"\n=> Summary: {total_questions} Questions total, {total_marks} Marks total")
    else:
        print("ERROR:", data.get('message'))
