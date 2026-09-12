import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
rrb_dir = os.path.dirname(script_dir)
if rrb_dir not in sys.path:
    sys.path.insert(0, rrb_dir)

import app

sections_list = [
    "Section A — Multiple Choice Questions (MCQ): 3 questions, 1 mark each",
    "Section B — Very Short Answer: 2 questions, 2 marks each",
    "Section C — Short Answer: 2 questions, 3 marks each",
    "Section D — Long Answer (Coding & Output): 2 questions, 5 marks each"
]

sec_text = "\n".join(sections_list)

prompt = f"""You are an expert ICSE question paper setter. Generate a complete descriptive question paper for Class 10 Computer Applications (Topic: 4. Constructors).

SECTIONS TO GENERATE (YOU MUST GENERATE ALL OF THESE SECTIONS - DO NOT STOP AFTER SECTION A):
{sec_text}

MANDATORY RULES:
1. You MUST generate ALL 4 sections in the 'sections' JSON array: Section A, Section B, Section C, and Section D.
2. In Section D (Long Answer), include Java class and constructor coding questions with exact indentation in ```java ``` code blocks.
3. Every section must have its exact requested question count.

OUTPUT JSON FORMAT:
{{
  "sections": [
    {{
      "section_label": "Section A",
      "section_title": "Multiple Choice Questions",
      "marks_per_question": 1,
      "instruction": "Choose the correct option.",
      "questions": [
        {{
          "number": 1,
          "question": "Question text...",
          "options": ["(a)...", "(b)...", "(c)...", "(d)..."],
          "sub_questions": [],
          "needs_image": false,
          "image_category": "none",
          "circuit_type": null,
          "params": {{}},
          "compound_name": null,
          "smiles": null,
          "image_keywords": []
        }}
      ]
    }},
    {{
      "section_label": "Section B",
      "section_title": "Very Short Answer",
      "marks_per_question": 2,
      "instruction": "Answer in 1-2 lines.",
      "questions": [...]
    }},
    {{
      "section_label": "Section C",
      "section_title": "Short Answer",
      "marks_per_question": 3,
      "instruction": "Explain with syntax and code snippets.",
      "questions": [...]
    }},
    {{
      "section_label": "Section D",
      "section_title": "Long Answer / Java Programs",
      "marks_per_question": 5,
      "instruction": "Write complete Java programs with constructors.",
      "questions": [...]
    }}
  ]
}}
Return ONLY valid JSON.
"""

result, provider, err = app.generate_ai_content(prompt, timeout=90)
if result:
    raw = result['candidates'][0]['content']['parts'][0]['text']
    parsed = app.safe_json_loads(raw)
    secs = parsed.get('sections', [])
    print(f"SUCCESS ({provider})! Total sections generated: {len(secs)}")
    for s in secs:
        print(f"  • {s.get('section_label')}: {s.get('section_title')} -> {len(s.get('questions', []))} questions")
        for q in s.get('questions', []):
            print(f"    Q{q.get('number')}. {q.get('question')[:65]}...")
else:
    print("Failed:", err)
