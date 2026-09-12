import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import image_handler

test_paper = {
    'sections': [
        {
            'section_label': 'Section A',
            'section_title': 'Programming & Output Questions (Python & Java)',
            'marks_per_question': 3,
            'instruction': 'Answer all coding questions with proper syntax and indentation.',
            'questions': [
                {
                    'number': 1,
                    'question': 'Predict the output of the following Python code snippet:\n```python\ndef calculate_series(n):\n    total = 0\n    for i in range(1, n + 1):\n        if i % 2 == 0:\n            total += i * 2\n        else:\n            total += i\n    return total\n\nresult = calculate_series(5)\nprint("Final Result =", result)\n```',
                    'options': [
                        '(a) Final Result = 15',
                        '(b) Final Result = 21',
                        '(c) Final Result = 25',
                        '(d) Final Result = 30'
                    ],
                    'sub_questions': []
                },
                {
                    'number': 2,
                    'question': 'Write a function in Python `count_vowels(text)` that accepts a string parameter and returns the total count of vowels (case-insensitive).',
                    'options': [],
                    'sub_questions': [
                        'Show sample test case with "Computer Science".',
                        'Write time complexity of your approach.'
                    ]
                }
            ]
        }
    ]
}

# Test 1: With Specific Topic
meta_with_topic = {
    'school_name': 'DELHI PUBLIC ACADEMY',
    'school_address': 'Sector 12, Institutional Area, New Delhi',
    'academic_session': '2024-2025',
    'exam_type': 'Pre-Board Examination',
    'class': '12',
    'subject': 'Computer Science',
    'topic': 'Python Functions and Loops',
    'full_syllabus': False,
    'duration': '3 Hours',
    'max_marks': '70',
    'teacher_name': 'Gaurav Shukla'
}

out_dir = os.path.join(script_dir, 'test_output')
os.makedirs(out_dir, exist_ok=True)
docx_path1 = os.path.join(out_dir, 'test_cs_paper_topic.docx')
image_handler.create_paper_docx(test_paper, meta_with_topic, docx_path1)
print(f"[PASS] DOCX (With Topic) Generated at: {docx_path1} ({os.path.getsize(docx_path1)} bytes)")

# Test 2: Full Syllabus (Topic must be omitted!)
meta_full_syl = dict(meta_with_topic)
meta_full_syl['topic'] = 'Full Syllabus'
meta_full_syl['full_syllabus'] = True

docx_path2 = os.path.join(out_dir, 'test_cs_paper_full_syllabus.docx')
image_handler.create_paper_docx(test_paper, meta_full_syl, docx_path2)
print(f"[PASS] DOCX (Full Syllabus - Topic Omitted) Generated at: {docx_path2} ({os.path.getsize(docx_path2)} bytes)")
