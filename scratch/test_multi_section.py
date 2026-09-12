import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import app

with app.app.test_request_context('/descriptive_paper', method='POST', data={
    'board': 'ICSE',
    'class': '10',
    'subject': 'Computer Application',
    'topics': '4. Constructors',
    'full_syllabus': 'no',
    'exam_type': 'Unit Test',
    'duration': '1 Hour',
    'max_marks': '25',
    'mcq_count': '3',
    'mcq_marks': '1',
    'vs_count': '2',
    'vs_marks': '2',
    'sh_count': '2',
    'sh_marks': '3',
    'lg_count': '2',
    'lg_marks': '5'
}):
    from flask import session
    session['admin_logged_in'] = True
    resp = app.descriptive_paper()

    raw_text = resp.get_data(as_text=True)
    import json
    data = json.loads(raw_text)
    print('Status:', data.get('status'))
    print('Provider:', data.get('provider'))
    if data.get('status') == 'success':
        sections = data.get('paper', {}).get('sections', [])
        print('Total Sections generated:', len(sections))
        for sec in sections:
            print(f"  • {sec.get('section_label')}: {sec.get('section_title')} -> {len(sec.get('questions', []))} Questions")
            for q in sec.get('questions', []):
                print(f"    Q{q.get('number')}. {q.get('question')[:60]}...")
    else:
        print('Error Message:', data.get('message'))


