import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import app

test_ai_output = '''```json
{
  "sections": [
    {
      "section_label": "Section A",
      "section_title": "Multiple Choice Questions",
      "questions": [
        {"number": 1, "question": "What is a constructor?", "options": ["(a) Method", "(b) Special method"]}
      ]
    },
    {
      "section_label": "Section D",
      "section_title": "Long Answer",
      "questions": [
        {
          "number": 2,
          "question": "Predict the output of the following Java snippet:\\n```java\\nclass Demo {\\n    Demo() {\\n        System.out.println(\\"Hello\\");\\n    }\\n}\\n```"
        }
      ]
    }
  ]
}
```'''

# Old buggy logic:
raw_old = test_ai_output.strip()
if '```' in raw_old:
    for part in raw_old.split('```'):
        part = part.strip()
        if part.startswith('json'): part = part[4:].strip()
        if part.startswith('{'): raw_old = part; break

print("OLD BUGGY RESULT SECTIONS COUNT:", len((app.safe_json_loads(raw_old) or {}).get('sections', [])))

# New robust logic:
raw_new = test_ai_output.strip()
if raw_new.startswith('```'):
    first_line_end = raw_new.find('\n')
    if first_line_end != -1:
        raw_new = raw_new[first_line_end+1:].strip()
    else:
        raw_new = raw_new.lstrip('`').strip()
        if raw_new.startswith('json'):
            raw_new = raw_new[4:].strip()
if raw_new.endswith('```'):
    raw_new = raw_new[:-3].strip()

print("NEW FIXED RESULT SECTIONS COUNT:", len((app.safe_json_loads(raw_new) or {}).get('sections', [])))
