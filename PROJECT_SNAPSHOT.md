# PROJECT SNAPSHOT: RRB CBT (Computer Based Test) System

> **Document Version:** 1.0 (Reflecting Codebase v1.13 Enhanced Edition)  
> **Target Audience:** Developers, System Administrators, AI Assistants  
> **Source Base:** `RRB_cbt-v1.10/RRB_v110`

---

## 1. PROJECT OVERVIEW

### Core Purpose
**RRB CBT** is an on-premise, web-based examination management and Computer Based Test (CBT) platform tailored for Indian school education environments (specifically RRB Group of Schools, originally developed by Gaurav Shukla). The platform automates the complete school examination lifecycle: from multi-provider AI-powered question paper generation and NCERT-aligned curriculum indexing to local-area-network (LAN) live exam delivery, anti-cheat client protection, real-time exam monitoring, automated multi-subject scoring with negative marking, student progress analytics, and print-ready Word (.docx)/PDF export.

### Tech Stack
| Tier | Technology / Library | Role & Details |
|---|---|---|
| **Backend Core** | **Python 3.9+** & **Flask 3.0.0** | Monolithic WSGI web service powering routes, session control, and template rendering (`app.py`, 9,500+ LOC). |
| **Database** | **SQLite 3** (`database.db`) | Single-file embedded relational DB; configured with `PRAGMA busy_timeout = 30000` and `PRAGMA journal_mode = DELETE` for concurrent transaction stability. |
| **Security & Rate Limiting** | **Werkzeug 3.0.1**, **Flask-Limiter 3.5.0**, **PyJWT 2.8.0** | SHA-256 password hashing, memory-based IP rate limiting (login & AI routes), and HS256 JWT tokens for REST APIs. |
| **Async Task Processing** | **Redis 5.0.1** + **RQ 1.15.1** | Background job management for heavy AI paper generation; uses `SimpleWorker` on Windows with automatic fallback to Python `threading.Thread` queue when Redis is offline. |
| **Document Generation** | **python-docx (>=1.1.0)**, **WeasyPrint 62.3**, **openpyxl 3.1.2** | Generation of native Word `.docx` documents (circuits, tables, callouts), printable HTML-to-PDF reports (via patched pydyf 0.12+), and Excel `.xlsx` ranking exports. |
| **Scientific & Diagramming** | **Schemdraw (>=0.19)**, **RDKit (>=2023.9.0)**, **PyMuPDF / fitz (>=1.23.0)** | Programmatic electric circuit schematics, 2D chemical structure depictions from SMILES, and NCERT PDF diagram extraction. |
| **External AI Providers** | **Google Gemini**, **DeepSeek**, **OpenAI**, **Claude**, **Hugging Face** | Multi-engine fallback pipeline for educational test synthesis and vision-based diagram extraction. |
| **Frontend UI** | **Jinja2**, **Vanilla CSS**, **Vanilla JS**, **MathJax v3**, **Chart.js** | Responsive server-rendered templates, LaTeX formula typesetting (`tex-svg.js` + `mhchem`), FontAwesome 6, Google Fonts (Poppins, Noto Sans Devanagari for bilingual Hindi/English papers), and Chart.js analytics. |

### Deployment Environment
- **Operating System:** Windows (primary production host).
- **Execution Workflow:** Managed via Windows batch scripts (`RRB_CBT_Manager.bat` menu orchestrator and `run.bat`). Runs `python app.py` on `0.0.0.0:5000` alongside optional background worker `python rq_worker.py`.
- **LAN / Network Architecture:** Deployed strictly on a local school Wi-Fi router or Ethernet LAN. Students and teachers access the server via the host machine's private IPv4 address (e.g., `http://192.168.1.X:5000/`). Exam delivery is designed to operate completely offline without external internet dependency, whereas AI question generation and Wikimedia image harvesting require an active internet connection on the host server.

---

## 2. USER ROLES & ACCESS CONTROL

The system defines **three primary actors** in code (`Admin`, `Teacher`, and `Student`), with granular role-based access decorators in `app.py`:

```
                    ┌────────────────────────┐
                    │      RRB CBT Platform  │
                    └───────────┬────────────┘
         ┌──────────────────────┼──────────────────────┐
         ▼                      ▼                      ▼
 ┌───────────────┐      ┌───────────────┐      ┌───────────────┐
 │     ADMIN     │      │    TEACHER    │      │    STUDENT    │
 └───────┬───────┘      └───────┬───────┘      └───────┬───────┘
         │                      │                      │
         │                      ├─ Class Teacher       │
         │                      └─ Subject Teacher     │
```

*(Note: While real-world school hierarchies include a Principal, **no dedicated "Principal" role exists in the code**. Oversight and class-wide administrative tasks are handled by Admins and designated Class Teachers).*

### Role 1: Administrator (`Admin`)
- **Authentication & Decorator:** 
  - Login via `/admin/login` (validated against environment variable `ADMIN_PASSWORD`, defaulting to `admin123`).
  - Enforced by `@admin_required` decorator (checks `session['admin_logged_in']`; returns HTTP 401 JSON for AJAX/APIs or redirects browser to `/admin/login`).
- **Permissions & Capabilities:**
  - **Exam Master Control:** Start/stop active exam sessions, set exam duration, toggle negative marking (with configurable penalty deduction, e.g., 0.33), and auto-submit expired sessions (`/admin/start_exam`, `/admin/stop_exam`, `auto_submit_expired_exams()`).
  - **Teacher Management:** Add/edit/delete teacher accounts, assign classes, sections, and subjects, designate Class Teachers, and perform bulk teacher imports via CSV (`/admin/teachers`, `/admin/teacher/add`, `/admin/teacher/upload_csv`).
  - **Question Bank Control:** Full CRUD on questions, bulk CSV import, wipe questions by class/subject, view print previews (`/admin/questions`, `/admin/upload_csv`, `/admin/generate_question_paper`).
  - **Live Monitoring & Evaluation:** Live real-time dashboard tracking connected student IP addresses, test status (`Not Started`, `In Progress`, `Submitted`), automated scoring evaluation, and Excel export (`/admin/monitoring`, `/admin/evaluate`, `/admin/export/results`).
  - **Security & Reattempts:** Review and approve/reject student exam reattempt requests (`/admin/reattempt_requests`, `handle_reattempt_request()`), unlock student class bindings (`student_class_lock`), and inspect system security events via audit logs (`/admin/audit_logs`).
  - **System Settings:** Configure school name, academic session, address, and upload institution branding logo (`/admin/settings`).

### Role 2: Teacher
- **Authentication & Decorator:**
  - Login via `/teacher/login` using registered mobile number and password (hashed with SHA-256 via `verify_password()`).
  - Enforced by `@teacher_required` decorator (checks `session['teacher_logged_in']`).
- **Sub-Designations & Scope:**
  - **Subject Teacher:** Scope restricted to classes and subjects assigned in `teacher_assignments`. Can only create tests, view questions, and access responses for their assigned subjects.
  - **Class Teacher (`is_class_teacher == 1`):** Holds elevated class-wide visibility for their assigned class and section. Can monitor student test scores, rank lists, and progress charts across *all* subjects in that class, and edit complete student profiles.
- **Permissions & Capabilities:**
  - **AI Test & Paper Generation:** Generate interactive CBT tests and formal multi-section descriptive papers (Unit Tests, Pre-Boards, Question Banks, DPP) using AI models or NCERT Exemplars (`/teacher/create_test`, `/descriptive_paper`, `/api/generate_paper_async`).
  - **Question Bank & Papers:** Upload subject-specific test papers via CSV, preview papers, and toggle active test papers (`/teacher/questions`, `/admin/teacher/upload_csv`, `/api/test_papers`).
  - **Student Monitoring & History:** View student profiles, upload student pictures, track performance analytics via Chart.js, download individual PDF student report cards, and monitor live test-takers (`/teacher/students`, `/teacher/student/<id>`, `/teacher/monitoring`).
  - **Reattempt Approvals:** Review and approve or reject reattempt requests submitted by students in their assigned classes/subjects (`/teacher/reattempt_requests`, `teacher_handle_reattempt_request()`).
  - **Shared Paper Repository:** Download and share generated `.docx` and `.doc` exam papers within the classwise shared repository (`/api/shared_papers/list`).

### Role 3: Student
- **Authentication & Decorator:**
  - Registration/Login via `/student_login` requiring Name, Student ID/Roll No, Class, Section, Subject, and Test No.
  - Bound by `student_class_lock` table (locks a student to their registered class/section).
  - Enforced by `@student_required` decorator (checks `session['student_id']`).
- **Permissions & Capabilities:**
  - **Waiting Room:** Access `/waiting` while waiting for the administrator/teacher to officially start the exam window.
  - **CBT Exam Interface:** Access `/exam` to take timed online tests. Includes randomized question ordering and option-shuffling with referential option protection.
  - **Response Autosave:** Periodically autosaves answers via `/save_answer` (translates displayed letter back to canonical database option letter).
  - **Anti-Cheat Restrictions:** Client-side blocking of right-click, copy (`Ctrl+C`), paste (`Ctrl+V`), cut (`Ctrl+X`), select all (`Ctrl+A`), text selection, dragging, and full-screen tab-switch tracking (`exam.js`).
  - **Submission & Reattempt:** Submit final answers via `/submit_exam`, view completion screen (`/submitted`), check test scores/answer keys (`/results`), or apply for a reattempt (`/request_reattempt`).

### Secondary API Authentication: JSON Web Tokens (JWT)
- Implemented via `generate_jwt_token()` and `decode_jwt_token()`.
- API endpoints are protected using `@jwt_required(allowed_roles=['admin', 'teacher', 'student'])`.
- Expects `Authorization: Bearer <token>` header; decodes payload claims (`sub`, `role`, `name`, `exp`), populating `g.jwt_user` and `g.jwt_payload`.


---

## 3. CORE FEATURES & IMPLEMENTATION LOGIC

### 3.1 Question Bank Management
- **Implementation:** `app.py` (`questions_data()`, `upload_csv()`, `add_question()`, `update_question()`, `delete_question()`, `delete_questions_by_class_subject()`).
- **Logic:**
  - Questions are stored in the `questions` table with fields for English & Hindi bilingual text, 4 choices (`option_a` through `option_d`), `correct_answer`, image path, class, section, subject, chapter, and test number.
  - Supports manual web-form entry or bulk CSV uploads.
  - `parse_question_csv_row()` and `_normalize_correct_answer()` automatically sanitize messy user inputs, mapping variants like `"Option A"`, `"A"`, `"1"`, or exact answer text into canonical `option_a` format.

### 3.2 Question & Option Shuffling Logic (with Referential Protection)
- **Implementation:** `app.py` (`_shuffle_options_safe()`, `_REFERENTIAL_OPT_RE`, `student_login()`, `exam()`, `get_questions()`, `save_answer()`).
- **Logic:**
  - **Question Order:** For each student, the question IDs (`qids`) are shuffled randomly via `random.shuffle(qids)`.
  - **Option Shuffling Bug Fix:** Historically, random shuffling moved referential options like *"Both a and b"*, *"None of these"*, or *"All of the above"* to invalid slots (e.g., *"Both a and b"* appearing at slot A).
  - **The Fix:** `_shuffle_options_safe()` inspects the 4 options using regex `_REFERENTIAL_OPT_RE`. If an option references letters or positions (`both`, `all of the above`, `none of these`, `except`, `only`, `(a)`, `(b)`), it is **pinned** to its original slot. Only the remaining normal options are shuffled.
  - **Storage & Mapping:** The final mapping is stored in `shuffled_questions.option_order` (e.g., `'BACD'`).
  - **Answer Key Integrity:** `save_answer()` reads `option_order`, translating the displayed selection back to the canonical database option letter before writing to `responses`. Pinned slots act as identity mappings, preserving 100% answer-key correctness.

### 3.3 AI-Based Question Generation
- **Implementation:** `app.py` (`generate_ai_content()`, `call_gemini_generate_content()`, `_call_deepseek_api()`, `_call_openai_api()`, `_call_claude_api()`, `_call_huggingface_api()`), grounded by `curriculum_data.py` (35,000+ LOC syllabus repository).
- **Logic:**
  - Multi-provider fallback cascade:
    1. **Google Gemini** (models: `gemini-3.6-flash`, `gemini-3.5-flash-lite`, `gemini-2.5-pro` with automatic exponential backoff on HTTP 429 quota exhaustion).
    2. **DeepSeek AI** (`deepseek-chat`).
    3. **OpenAI ChatGPT** (`gpt-4o-mini`, `gpt-4o`, `gpt-3.5-turbo`).
    4. **Anthropic Claude** (`claude-3-5-haiku-20241022`, `claude-3-haiku-20240307`).
    5. **Hugging Face Inference API** (`Qwen/Qwen2.5-7B-Instruct`, `meta-llama/Llama-3.1-8B-Instruct`).
  - Prompts are seeded with rationalized CBSE/ICSE curriculum guidelines from `curriculum_data.py`. Responses are enforced as structured JSON arrays with schema sanitization and markdown fence stripping.

### 3.4 Image & Diagram Pipeline
- **Implementation:** `image_handler.py` (1,600+ LOC) and `exemplar_extractor.py` (750+ LOC).
- **Sources & Strategies:**
  1. **Schemdraw Circuit Engine:** Programmatically generates circuit schematics (series/parallel resistors, RLC circuits, Wheatstone bridges, logic gate diagrams) directly to PNG.
  2. **Local NCERT Image Bank:** Pre-indexed repository (`image_bank_index.json`) containing curated biology, physics, geography, and chemistry diagrams matched via fuzzy keyword searching.
  3. **Wikimedia Commons API:** On-demand downloader that queries Wikimedia Commons for CC-BY / CC-BY-SA / Public Domain educational diagrams (validated by `_is_allowed_wikimedia_license()`).
  4. **Chemical Structure Renderer:** Chemical names are matched against `NCERT_COMPOUND_SMILES`. RDKit generates 2D depictions, with fallback to the PubChem PUG REST API.
  5. **NCERT Exemplar PDF Extractor:** `exemplar_extractor.py` uses PyMuPDF (`fitz`) to crop images from official NCERT Exemplar PDFs and pairs them with questions using Gemini Vision (`_call_gemini_vision_page()`).
  6. **Callout Placeholders:** When an image cannot be located automatically, `insert_placeholder()` injects a styled yellow callout box into Word/HTML papers so teachers can paste or sketch figures manually.

### 3.5 Live Exam Mode & Anti-Cheat System
- **Implementation:** `app.py` (`start_exam()`, `stop_exam()`, `monitoring()`, `monitoring_data()`, `auto_submit_expired_exams()`) and `static/js/exam.js`.
- **Logic:**
  - Admin/Teacher activates the exam for a class/subject. Students in `/waiting` poll `/api/exam_status` and are automatically transitioned to `/exam`.
  - Timer and duration are enforced server-side. If the clock expires, `auto_submit_expired_exams()` automatically forces status to `Submitted`.
  - Client-side security in `exam.js`: Blocks context menu, clipboard events (`copy`, `paste`, `cut`), selection (`selectstart`), and full-window blur tracking.

### 3.6 Test Paper Generation & Document Exports
- **Implementation:** `app.py` (`descriptive_paper()`, `save_paper_to_repository()`, `create_paper_docx()`, `process_ai_paper_job()`, `enqueue_ai_paper_job()`), `image_handler.py`.
- **Supported Formats:**
  - **Native Word (`.docx`):** Built via `python-docx` with institutional headers, metadata tables, embedded high-res diagrams, formatted LaTeX formulas, and answer keys.
  - **Printable HTML / Word (`.doc`):** Rendered via Jinja2 templates with MathJax typesetting.
  - **Print / Download PDF:** Built using WeasyPrint with custom pydyf 0.12+ monkey-patches (`_patched_pydyf_pdf_init`, `_patched_stream_transform`).
  - **Excel (`.xlsx`):** Built using `openpyxl` for student marksheets and rank lists.
- **Classwise Shared Repository:** Generated papers are saved into:
  `paper/Class_<N>/Section_<X>/<class>_<section>_<subject>_<paper_type>_<TeacherName>.docx`
  All teachers can browse and download papers across classes for collaborative reuse (`/api/shared_papers/list`).

### 3.7 Asynchronous Paper Generation Queue
- **Implementation:** `app.py` (lines 7957–8180), `rq_worker.py`.
- **Logic:**
  - When Redis is running (`REDIS_HOST:6379`), paper generation tasks are enqueued to RQ (`rq_paper_queue.enqueue()`), processed by `rq_worker.py` (using `SimpleWorker` on Windows).
  - If Redis is unavailable, the system automatically falls back to an internal daemon thread worker (`ai_paper_thread_worker()`).
  - Clients poll `/api/paper_job_status/<job_id>`. Upon completion, a notification record is inserted into `teacher_notifications`.

### 3.8 Approval Workflows (Reattempt Security)
- **Implementation:** `app.py` (`request_reattempt()`, `admin_reattempt_requests()`, `teacher_reattempt_requests()`, `handle_reattempt_request()`, `teacher_handle_reattempt_request()`).
- **Logic:**
  - If a student gets disconnected or accidentally submits, they cannot log in again due to `student_class_lock` and student status `Submitted`.
  - The student submits a reattempt request via `/request_reattempt`. Status becomes `pending` in `reattempt_requests`.
  - The Admin or the student's assigned Teacher reviews the request in their dashboard.
  - **On Approval:** The student status is reset to `'Not Started'`, `exam_started_at` is cleared, entries in `shuffled_questions` are purged (allowing a clean reshuffle), previous scores are archived/cleared, and the student can re-enter `/exam`.
  - **On Rejection:** The status is updated to `rejected` with an admin/teacher remark.


---

## 4. DATABASE SCHEMA

The SQLite database (`database.db`) consists of **21 tables**. Below is the complete schema definition:

### 1. `students`
Represents student test-takers, session state, and biographical profiles.
- `student_id` (TEXT, PRIMARY KEY): Unique student identifier / roll number.
- `name` (TEXT): Full student name.
- `class` (TEXT): Enrolled class (e.g. `10`, `12`).
- `section` (TEXT): Section identifier (e.g. `A`, `B`).
- `subject` (TEXT): Current test subject.
- `test_no` (TEXT): Active test identifier.
- `ip` (TEXT): Client IPv4 address.
- `status` (TEXT): `Not Started`, `In Progress`, or `Submitted`.
- `exam_started_at` (TIMESTAMP): Time when student started the exam.
- `admission_no` (TEXT): School admission number.
- `dob` (TEXT): Date of birth.
- `house` (TEXT): School house assignment.
- `parents_name` (TEXT): Parent/guardian names.
- `address` (TEXT): Residential address.
- `picture` (TEXT): Path to profile picture.
- `system_name` (TEXT): Computer/terminal hostname.

### 2. `questions`
Represents individual exam questions in the central question bank.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Question ID.
- `class` (TEXT): Associated class.
- `section` (TEXT): Associated section (optional).
- `subject` (TEXT): Subject name.
- `chapter` (TEXT): Chapter or topic title.
- `test_no` (TEXT): Test paper identifier.
- `question_type` (TEXT): `MCQ`, `Assertion-Reason`, `Short`, `Long`, `Case Study`.
- `question` (TEXT): Question text (English / primary).
- `question_hi` (TEXT): Question text in Hindi (for bilingual papers).
- `option_a`, `option_b`, `option_c`, `option_d` (TEXT): Options in English.
- `option_a_hi`, `option_b_hi`, `option_c_hi`, `option_d_hi` (TEXT): Options in Hindi.
- `correct_answer` (TEXT): Canonical answer (`option_a`, `option_b`, `option_c`, `option_d`).
- `image_path` (TEXT): Path to associated diagram or `__smiles__` chemical string.

### 3. `shuffled_questions`
Maps randomized question orders and option permutations per student.
- `student_id` (TEXT): References `students(student_id)`.
- `question_id` (INTEGER): References `questions(id)`.
- `shuffled_index` (INTEGER): Display sequence position (0, 1, 2...).
- `option_order` (TEXT): 4-character permutation string (e.g. `'CABD'`).
- *PRIMARY KEY:* `(student_id, question_id)`

### 4. `responses`
Records live student answers.
- `student_id` (TEXT): References `students(student_id)`.
- `question_id` (INTEGER): References `questions(id)`.
- `selected_option` (TEXT): Canonical original option letter (`A`, `B`, `C`, `D`).
- `created_at` (TIMESTAMP): Timestamp of answer selection.
- *PRIMARY KEY:* `(student_id, question_id)`

### 5. `results`
Historical test scores and percentage records.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Result ID.
- `student_id` (TEXT): Student identifier.
- `name` (TEXT): Student name.
- `class` (TEXT): Class.
- `section` (TEXT): Section.
- `subject` (TEXT): Subject.
- `score` (INTEGER / REAL): Final score after negative marking deductions.
- `total_questions` (INTEGER): Total possible questions / marks.
- `percentage` (REAL): Calculated percentage score.
- `chapter` (TEXT): Test number or chapter name.
- `test_date` (TIMESTAMP): Exam date.

### 6. `exam_control`
Singleton table controlling global exam runtime parameters (row `id = 1`).
- `id` (INTEGER, PRIMARY KEY CHECK (id=1)): Fixed ID.
- `is_active` (INTEGER): `1` if exam is active, `0` if stopped.
- `start_time` (TIMESTAMP): Server timestamp when exam was launched.
- `duration` (INTEGER): Duration in minutes.
- `negative_marking` (INTEGER): `1` if enabled, `0` if disabled.
- `negative_value` (REAL): Penalty mark deduction per incorrect answer (e.g. `0.33`).

### 7. `settings`
Key-value configuration store for institutional branding.
- `key` (TEXT, PRIMARY KEY): Config name (`school_name`, `logo_path`, `school_address`, `academic_session`).
- `value` (TEXT): Stored value.

### 8. `teachers`
Registered teacher credentials and profiles.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Teacher ID.
- `name` (TEXT): Full teacher name.
- `mobile` (TEXT, UNIQUE): Login mobile number.
- `password` (TEXT): SHA-256 password hash.
- `email` (TEXT): Email address.
- `address` (TEXT): Residential address.
- `picture` (TEXT): Profile photo path.
- `status` (TEXT): `active` or `inactive`.
- `created_at` (TIMESTAMP): Account creation timestamp.

### 9. `teacher_assignments`
Class and subject assignments linking teachers to teaching scopes.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Assignment ID.
- `teacher_id` (INTEGER): Foreign Key references `teachers(id)`.
- `class` (TEXT): Assigned class.
- `section` (TEXT): Assigned section.
- `subject` (TEXT): Assigned subject.
- `is_class_teacher` (INTEGER): `1` if designated Class Teacher, `0` if Subject Teacher.

### 10. `test_papers`
Metadata for uploaded CSV test paper files.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Paper ID.
- `filename` (TEXT): Stored file name.
- `class` (TEXT): Class.
- `section` (TEXT): Section.
- `subject` (TEXT): Subject.
- `test_no` (TEXT): Test paper code.
- `paper_type` (TEXT): `Descriptive Paper`, `Question Bank`, `DPP`.
- `uploaded_by` (TEXT): Uploader username.
- `uploader_type` (TEXT): `admin` or `teacher`.
- `is_active` (INTEGER): `1` if available for examination, `0` if inactive.
- `question_count` (INTEGER): Number of parsed questions.
- `created_at` (TIMESTAMP): Upload timestamp.

### 11. `reattempt_requests`
Audit queue for student exam reattempts.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Request ID.
- `student_id` (TEXT): Requesting student.
- `class` (TEXT): Class.
- `subject` (TEXT): Subject.
- `status` (TEXT): `pending`, `approved`, or `rejected`.
- `admin_note` (TEXT): Feedback or justification from reviewer.
- `requested_at` (TIMESTAMP): Submission time.
- `reviewed_at` (TIMESTAMP): Review time.

### 12. `student_class_lock`
Security binding preventing students from switching classes/sections.
- `student_id` (TEXT, PRIMARY KEY): Student ID.
- `class` (TEXT): Locked class.
- `section` (TEXT): Locked section.
- `locked_at` (TIMESTAMP): Lock timestamp.

### 13. `bulletins`
Announcements displayed on student login and dashboard pages.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Bulletin ID.
- `title` (TEXT): Notice headline.
- `content` (TEXT): Notice body.
- `posted_by` (TEXT): Author name.
- `poster_type` (TEXT): `admin` or `teacher`.
- `target_class` (TEXT): Specific class or empty for all.
- `is_active` (INTEGER): `1` if visible, `0` if hidden.
- `created_at` (TIMESTAMP): Creation date.

### 14. `test_generation_history`
Audit log of all AI-generated question papers.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): History ID.
- `teacher_id` (TEXT): Author teacher identifier.
- `class`, `section`, `subject`, `chapter`, `test_no` (TEXT): Paper parameters.
- `paper_type` (TEXT): Paper style.
- `output_mode` (TEXT): `cbt` or `descriptive`.
- `total_questions`, `mcq_count`, `assertion_count`, `very_short_count`, `short_count`, `long_count`, `case_study_count` (INTEGER): Breakdown of question types.
- `remark` (TEXT): Generation notes.
- `created_at` (TIMESTAMP): Timestamp.

### 15. `mcq_test_history`
Quick-reference table tracking MCQ test generation across teachers.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Record ID.
- `teacher_id` (INTEGER): Teacher ID.
- `teacher_name` (TEXT): Teacher name.
- `class`, `section`, `subject`, `test_no` (TEXT): Test specifications.
- `question_count` (INTEGER): Count of questions.
- `created_at` (TIMESTAMP): Generation timestamp.

### 16. `combined_tests` & 17. `combined_test_subjects`
Entities supporting multi-subject composite examinations (e.g. Physics + Chemistry + Math).
- `combined_tests`: `id`, `test_no`, `class`, `section`, `title`, `created_by`, `creator_type`, `is_active`, `created_at`.
- `combined_test_subjects`: `id`, `combined_test_id` (FK references `combined_tests(id)` ON DELETE CASCADE), `subject`, `question_count`.

### 18. `scheduled_tests`
Supports automated test windows and scheduled duration control.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Schedule ID.
- `class`, `section`, `subject`, `test_no` (TEXT): Exam parameters.
- `scheduled_date` (TEXT): Date string (`YYYY-MM-DD`).
- `start_time`, `end_time` (TEXT): Time strings (`HH:MM`).
- `duration_minutes` (INTEGER): Exam length.
- `status` (TEXT): `scheduled`, `running`, `completed`, `cancelled`.
- `created_by` (TEXT): Admin username.
- `created_at` (TIMESTAMP): Creation timestamp.

### 19. `teacher_notifications`
In-app alert bell notifications for async AI paper generation completion.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Notification ID.
- `teacher_id` (INTEGER): Recipient teacher.
- `title` (TEXT): Notification title.
- `message` (TEXT): Detailed message.
- `link` (TEXT): Direct URL to generated paper.
- `is_read` (INTEGER): `0` (unread) or `1` (read).
- `created_at` (TIMESTAMP): Notification time.

### 20. `audit_logs`
Security and administrative compliance logging.
- `id` (INTEGER, PRIMARY KEY AUTOINCREMENT): Log ID.
- `user_type` (TEXT): `admin`, `teacher`, `student`.
- `user_id` (TEXT): User identifier.
- `action` (TEXT): Action performed (e.g. `LOGIN`, `DELETE_QUESTION`, `APPROVE_REATTEMPT`).
- `target_table` (TEXT): Affected table name.
- `target_id` (TEXT): Affected record ID.
- `ip_address` (TEXT): Client IP.
- `timestamp` (TIMESTAMP): Action timestamp.

### 21. `sqlite_sequence`
Internal SQLite system table managing `AUTOINCREMENT` keys.

### Key Entity Relationships
- **Teachers & Assignments:** `teachers.id` (1) ── (N) `teacher_assignments.teacher_id`
- **Composite Tests:** `combined_tests.id` (1) ── (N) `combined_test_subjects.combined_test_id`
- **Shuffled Exam Delivery:** `students.student_id` (1) ── (N) `shuffled_questions.student_id` (N) ── (1) `questions.id`
- **Student Responses:** `students.student_id` (1) ── (N) `responses.student_id` (N) ── (1) `questions.id`
- **Reattempt Queue:** `students.student_id` (1) ── (N) `reattempt_requests.student_id`
- **Student Class Lock:** `students.student_id` (1) ── (1) `student_class_lock.student_id`


---

## 5. FILE & FOLDER STRUCTURE

```
RRB_cbt-v1.10/RRB_v110/
├── app.py                      # Core monolith (9,546 lines): all HTTP routing, session auth, exam execution, DB migrations
├── curriculum_data.py          # Master curriculum repository (35,706 lines): CBSE/ICSE/State syllabus for Classes 1-12
├── image_handler.py            # Scientific image engine (1,613 lines): Schemdraw circuits, RDKit chemistry, Wikimedia, docx builder
├── exemplar_extractor.py       # NCERT Exemplar extraction (752 lines): PyMuPDF PDF parsing & Gemini Vision diagram OCR
├── validators.py               # Request schema validation engine (137 lines): rule dictionary validator & decorators
├── rq_worker.py                # Background job queue worker (37 lines): Redis RQ SimpleWorker for asynchronous paper generation
├── push_to_github.py           # GitHub sync automation script (238 lines): API-based repository creation & multi-chunk file push
│
├── database.db                 # Primary SQLite 3 database file containing all system tables
├── requirements.txt            # Python dependencies (Flask, weasyprint, openpyxl, redis, rq, schemdraw, rdkit, PyMuPDF, etc.)
├── RRB_CBT_Manager.bat         # Comprehensive Windows CLI manager (run server, start worker, backup DB, manage dependencies)
├── Push_Version_To_GitHub.bat  # Automated wrapper script for version staging to GitHub
├── apikey.env / .api_key       # Local configuration files storing API keys for Gemini, Claude, DeepSeek, OpenAI
│
├── .agents/rules/              # System architecture and workflow specification rules
│   ├── concurrency_architecture.md   # Concurrency limits, API quotas, and async task queue patterns
│   ├── mcq_test_history.md           # MCQ test history logging and management rules
│   ├── progress_modal_animation.md   # Frontend generation modal animations and progress milestones
│   ├── shared_paper_repository.md    # Classwise hierarchical storage folder and file naming rules
│   └── teacher_assignment_rules.md   # Access scope and Class Teacher vs Subject Teacher rules
│
├── paper/                      # Shared Paper Repository (Hierarchical)
│   └── Class_<N>/
│       └── Section_<X>/        # <class>_<section>_<subject>_<paper_type>_<TeacherName>.docx / .doc
│
├── exemplar_bank/              # NCERT Exemplar assets
│   ├── images/                 # Cropped figures and diagrams extracted from exemplar PDFs
│   └── exemplar_bank_index.json# Index mapping subjects, chapters, and question numbers to extracted images
├── image_bank_index.json       # Metadata index for local NCERT diagrams (biology, chemistry, physics, geography)
│
├── static/                     # Web assets served to the client
│   ├── css/style.css           # Primary application stylesheet
│   ├── js/exam.js              # Client-side exam runner: autosave, countdown timer, anti-cheat blocking
│   ├── js/multimodal.js        # Multimodal media integration helper
│   ├── images/ncert/           # Local educational diagram PNGs (human heart, refraction, logic gates, etc.)
│   ├── student_profile_pic/    # Uploaded student identification photos
│   └── Teacher_profile_picture/# Uploaded teacher identification photos
│
├── templates/                  # Jinja2 HTML UI templates (40 files)
│   ├── index.html              # Student login and exam access portal
│   ├── waiting.html            # Exam waiting room for registered students
│   ├── exam.html               # Live timed CBT exam interface
│   ├── submitted.html          # Exam completion and reattempt request submission page
│   ├── results.html            # Student exam scorecard and answer review
│   ├── admin_login.html        # Master administrator authentication screen
│   ├── admin_dashboard.html    # Main administrative dashboard
│   ├── monitoring.html         # Real-time live exam monitoring console
│   ├── teacher_login.html      # Teacher portal login screen
│   ├── teacher_dashboard.html  # Teacher workspace and subject dashboard
│   ├── descriptive_paper.html  # Multi-section formal test paper generator
│   └── ...                     # Additional administrative, teacher, and reporting templates
│
├── exports/                    # Directory for generated Excel (.xlsx) result spreadsheets
├── backups/                    # Timestamped SQLite database backup copies
├── logs/                       # Server runtime logs (`logs/app.log`)
└── scratch/                    # Automated unit, integration, and verification test scripts
```

---

## 6. VERSION HISTORY & RECENT CHANGELOG

### Version 1.13 ("Enhanced Edition")
- **Anti-Cheat Hardening:** Injected clipboard and event blocking into `exam.js` (disables right-click, `Ctrl+C`, `Ctrl+V`, `Ctrl+X`, `Ctrl+A`, text dragging, and blur tracking).
- **Teacher Portal:** Full mobile + password authentication, teacher assignments, and dedicated dashboard (`/teacher`).
- **AI Paper Generation:** Integration of multi-provider LLMs for custom test synthesis with Claude / Gemini.
- **Student Profile Management:** Comprehensive student records (admission no, DOB, house, parents, photo) with Class Teacher editing privileges.
- **Progress Tracking:** Integration of Chart.js line charts tracking individual student performance across test dates.
- **Bulk Import:** CSV-based bulk import for teachers and questions.

### Recent Critical Fixes & Enhancements
- **Referential Option Shuffle Fix (`_shuffle_options_safe`):** Resolved critical exam bug where referential options (*"Both a and b"*, *"None of these"*, *"All of the above"*, *"Except a and b"*) became invalid when shuffled. The new logic pins referential options to their original slots while continuing to randomize standard options.
- **Negative Marking Calculation Fix (`BUG-003`):** Fixed evaluation logic in `evaluate()` to fetch negative marking settings first and compute deductions accurately before writing final scores to `results`.
- **Exam Isolation & Class-Lock (`BUG-001`):** Prevented cross-test collisions by isolating student submissions strictly by test ID and introducing `student_class_lock`.
- **AI Quota & Gemini Fallback Handling:** Added automatic candidate model fallbacks (`gemini-3.6-flash` -> `gemini-3.5-flash-lite` -> `gemini-2.5-pro`) and HTTP 429 exponential backoff.
- **Async Queue Integration:** Implemented Redis + RQ with automatic fallback to Python `threading.Thread` workers for high-concurrency test generation.
- **WeasyPrint 60+ / pydyf Compatibility:** Resolved PDF rendering incompatibilities in WeasyPrint 60+ stream transformation and pydyf 0.12+.

---

## 7. KNOWN ARCHITECTURAL CHARACTERISTICS & DESIGN NOTES

1. **Monolithic Architecture:** `app.py` contains over 9,500 lines encompassing routing, database queries, business logic, rendering, and API handlers. Care must be taken during edits to avoid breaking global variable states.
2. **SQLite Under High Concurrency:** SQLite is configured with `PRAGMA busy_timeout = 30000` and `PRAGMA journal_mode = DELETE`. While stable for classroom sizes (50–150 students), simultaneous mass exam submissions generate rapid database writes. The `_save_or_update_result()` helper handles upsert contention, but heavy write spikes should be monitored.
3. **Absence of a "Principal" Role:** The system implements access boundaries via Admin and Class Teachers. Requests mentioning "Principal approvals" in administrative contexts map to Admin-level approvals or Class Teacher reviews.
4. **Internet Dependency for AI / Wikimedia:** While student exam delivery operates 100% offline within a school LAN, AI generation routes and Wikimedia diagram queries require active external internet connectivity on the server host.
5. **Windows RQ Worker Constraints:** Standard Unix `rq worker` relies on `os.fork()`, which is unsupported on Windows. On Windows platforms, `rq_worker.py` must use `SimpleWorker` (or rely on the application's built-in thread fallback).

