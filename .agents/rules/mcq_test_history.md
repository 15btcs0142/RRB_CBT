# MCQ Test History Rules for All Teachers

## Table Schema (`mcq_test_history`)
- `id`: INTEGER PRIMARY KEY AUTOINCREMENT
- `teacher_id`: INTEGER
- `teacher_name`: TEXT
- `class`: TEXT
- `section`: TEXT
- `subject`: TEXT
- `test_no`: TEXT
- `question_count`: INTEGER
- `created_at`: TIMESTAMP DEFAULT CURRENT_TIMESTAMP

## Features & Controls
1. Log every MCQ test generation across all teachers.
2. Provide filterable table in Teacher Dashboard (Class, Section, Subject, Test No).
3. Provide quick actions: Preview Test, Download PDF, Export CSV, and Delete history record.
