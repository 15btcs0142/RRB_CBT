# RRB CBT Application - Quick Reference & Architecture Guide

## 🏗️ Application Architecture

```
RRB_CBT (Flask Web Application)
│
├── 🎓 Student Module
│   ├── Login & Registration
│   ├── Exam Interface
│   ├── Question Display (with shuffling)
│   ├── Answer Submission
│   └── Results & Reattempt Request
│
├── 🔐 Admin Module
│   ├── Admin Dashboard
│   ├── Question Bank Management
│   ├── Student Management
│   ├── Exam Control (Start/Stop)
│   ├── Real-time Monitoring
│   ├── Results & Export
│   └── Settings
│
├── 👨‍🏫 Teacher Module
│   ├── Teacher Dashboard
│   ├── Student Monitoring
│   ├── Student Profile Management
│   ├── Test History
│   └── Response Viewing
│
└── 🗄️ Database (SQLite)
    ├── Students Table
    ├── Questions Table
    ├── Responses Table
    ├── Results Table
    ├── Teachers Table
    ├── Exam Control
    └── Other Supporting Tables
```

---

## 🔄 Key Workflows

### 1️⃣ EXAM CREATION & MANAGEMENT

**Admin creates exam:**
```
Questions uploaded (CSV) → questions_data() → Questions stored in DB
         ↓
Admin starts exam → start_exam() → exam_control.is_active = 1
         ↓
Students see waiting page → waiting()
         ↓
Admin can monitor → monitoring_data()
         ↓
Admin stops exam → stop_exam() → exam_control.is_active = 0
```

**Question Upload Process:**
- Admin uploads CSV file → `upload_csv()`
- File is parsed and validated
- Questions inserted into database with class, subject, chapter info
- Questions have 4 options (A, B, C, D) and a correct answer

### 2️⃣ STUDENT EXAM WORKFLOW

```
┌─────────────────────────────────┐
│ Student visits home page (/)    │
└──────────────┬──────────────────┘
               │ (index())
               ↓
    ┌──────────────────────┐
    │ Select Class/Subject │
    │ Student ID           │
    └──────────────┬───────┘
                   │ (student_login())
                   ↓
          ┌────────────────────┐
          │ Check Class Lock   │
          │ (first time = lock)│
          └────────────┬───────┘
                       │
                       ↓
              ┌─────────────────┐
              │ Waiting Page    │ (waiting())
              │ Poll for exam   │
              │ start           │
              └────────┬────────┘
                       │ (check_exam_status())
                       ↓ [Exam starts]
              ┌─────────────────┐
              │ Load Exam Page  │ (exam())
              │ Questions       │ (get_questions())
              │ shuffled        │
              └────────┬────────┘
                       │
              ┌────────▼────────┐
              │ Answer Question │ (save_answer())
              │ → Next Question │ [Repeat]
              └────────┬────────┘
                       │
              ┌────────▼────────┐
              │ Submit Exam     │ (submit_exam())
              │ Calculate Score │
              └────────┬────────┘
                       │
              ┌────────▼──────────────┐
              │ Show Results          │ (submitted())
              │ Option to Reattempt   │
              └───────────────────────┘
```

### 3️⃣ SCORING SYSTEM

**Score Calculation:**
```python
score = number_of_correct_answers
if negative_marking_enabled:
    score -= (wrong_answers × negative_value)
percentage = (score / total_questions) × 100
```

**Results stored with:**
- Student ID & Name
- Class & Subject
- Score & Percentage
- Total Questions
- Test Date

### 4️⃣ QUESTION SHUFFLING

**Why shuffle?**
- Prevent cheating
- Same questions appear in different order for different students
- Each student gets different option orders for same questions

**Process:**
```
1. Get all questions for exam
2. Shuffle question order (Fisher-Yates algorithm)
3. For each question, shuffle option order
4. Store shuffled order in shuffled_questions table
5. When displaying, use shuffled order for THIS student
```

**Functions involved:**
- `shuffle_questions_for_student()` - Creates shuffled version
- `get_shuffled_questions()` - Retrieves shuffled questions
- `get_questions()` - Returns questions in shuffled order to frontend

---

## 📊 Database Relationships

```
students (1) ──→ (many) responses
                         ↓
                    questions (answered)

students (1) ──→ (many) results

students (1) ──→ (1) student_class_lock

teachers (1) ──→ (many) teacher_assignments
                         ↓
                    classes & subjects

exam_control (singleton) - Global exam state

shuffled_questions - Links students to their shuffled question orders

combined_tests (1) ──→ (many) combined_test_subjects
```

---

## 🔑 Key Features

### 1. Question Shuffling
- Questions appear in random order for each student
- Multiple choice options are also randomized
- Prevents students from using memorized positions

### 2. Negative Marking (Optional)
- Admin can enable/disable
- Wrong answers deduct points
- Formula: `score = correct - (wrong × negative_value)`
- Default: -0.33 per wrong answer

### 3. Reattempt Management
- Students can request reattempt after submission
- Admin can approve/reject requests
- Tracks all reattempt attempts

### 4. Class Lock
- Once a student takes an exam in Class X
- They can only take exams in that class
- Prevents students from switching classes
- Admin must manually unlock if needed

### 5. Combined Tests
- Multi-subject exams
- Questions from multiple subjects in one test
- Admin can set test numbers to map subjects

### 6. Real-time Monitoring
- Admin sees all students taking exam
- Knows who started, who's in progress, who submitted
- Can force-submit a student's exam
- Can view any student's responses

### 7. Exam Time Control
- Admin sets exam duration (default: 60 minutes)
- Countdown timer shows on student exam page
- Auto-submit when time expires
- `auto_submit_expired_exams()` runs automatically

---

## 🎨 Frontend Integration

### Key JavaScript Files
- **exam.js** - Handles exam page interactions
  - Timer countdown
  - Question navigation
  - Answer saving
  - Exam submission
  
- **multimodal.js** - Handles multi-subject exam display

### Key HTML Templates
- **exam.html** - Main exam interface
- **results.html** - Results display for students
- **admin_monitoring.html** - Real-time monitoring
- **teacher_monitoring.html** - Teacher's student monitoring

---

## 📁 File Structure

```
# RRB CBT Application - Quick Reference & Architecture Guide

## 🏗️ Application Architecture

```
RRB_CBT (Flask Web Application)
│
├── 🎓 Student Module
│   ├── Login & Registration
│   ├── Exam Interface & Timer Countdown
│   ├── Question Display (with shuffling & multi-subject support)
│   ├── Answer Submission & Real-time Auto-save
│   └── Results & Reattempt Requests
│
├── 🔐 Admin Module
│   ├── Admin Dashboard
│   ├── Question Bank Management (CSV Import & Manual Creation)
│   ├── Student & Teacher Management
│   ├── Exam Control (Start/Stop/Time Management)
│   ├── Real-time Monitoring & Student Progress Tracking
│   ├── Results Filter, Export (PDF/Excel) & Reports
│   └── Settings & School Branding
│
├── 👨‍🏫 Teacher Module
│   ├── Teacher Dashboard
│   ├── Student Monitoring
│   ├── Student Profile Management
│   ├── Test History & Class Reports
│   └── Response Viewing
│
├── 🚀 Versioning & Deployment Module
│   ├── GitHub Repo Sync (https://github.com/15btcs0142/RRB_CBT)
│   ├── Version Branching (v1.0, v1.1, v1.2, v1.3...)
│   ├── Git Tagging & Release Management
│   └── push_to_github.py & Push_Version_To_GitHub.bat
│
└── 🗄️ Database (SQLite)
    ├── Students Table
    ├── Questions Table
    ├── Responses Table
    ├── Results Table
    ├── Teachers Table
    ├── Exam Control & Settings Tables
    ├── Shuffled Questions Table
    └── Combined Tests Tables
```

---

## 🔄 Key Workflows

### 1️⃣ EXAM CREATION & MANAGEMENT

**Admin creates exam:**
```
Questions uploaded (CSV) → upload_csv() → Questions stored in DB
         ↓
Admin starts exam → start_exam() → exam_control.is_active = 1
         ↓
Students see waiting page → waiting()
         ↓
Admin can monitor → monitoring_data()
         ↓
Admin stops exam → stop_exam() → exam_control.is_active = 0
```

**Question Upload Process:**
- Admin uploads CSV file → `upload_csv()`
- File is parsed and validated
- Questions inserted into database with class, subject, chapter info
- Questions have 4 options (A, B, C, D) and a correct answer

---

### 2️⃣ STUDENT EXAM WORKFLOW

```
┌─────────────────────────────────┐
│ Student visits home page (/)    │
└──────────────┬──────────────────┘
               │ (index())
               ↓
    ┌──────────────────────┐
    │ Select Class/Subject │
    │ Student ID           │
    └──────────────┬───────┘
                   │ (student_login())
                   ↓
          ┌────────────────────┐
          │ Check Class Lock   │
          │ (first time = lock)│
          └────────────┬───────┘
                       │
                       ↓
              ┌─────────────────┐
              │ Waiting Page    │ (waiting())
              │ Poll for exam   │
              │ start           │
              └────────┬────────┘
                       │ (check_exam_status())
                       ↓ [Exam starts]
              ┌─────────────────┐
              │ Load Exam Page  │ (exam())
              │ Questions       │ (get_questions())
              │ shuffled        │
              └────────┬────────┘
                       │
              ┌────────▼────────┐
              │ Answer Question │ (save_answer())
              │ → Next Question │ [Repeat]
              └────────┬────────┘
                       │
              ┌────────▼────────┐
              │ Submit Exam     │ (submit_exam())
              │ Calculate Score │
              └────────┬────────┘
                       │
              ┌────────▼──────────────┐
              │ Show Results          │ (submitted())
              │ Option to Reattempt   │
              └───────────────────────┘
```

---

### 3️⃣ GITHUB VERSIONING & DEPLOYMENT WORKFLOW

```
┌──────────────────────────────────────────────────────────┐
│ Update local project code (app.py, templates, static)     │
└──────────────────────────┬───────────────────────────────┘
                           │
                           ↓
┌──────────────────────────────────────────────────────────┐
│ Run Push_Version_To_GitHub.bat (or push_to_github.py)    │
└──────────────────────────┬───────────────────────────────┘
                           │
                           ↓
┌──────────────────────────────────────────────────────────┐
│ Specify Version Tag (e.g. v1.1, v1.2, v1.3...)           │
└──────────────────────────┬───────────────────────────────┘
                           │
                           ↓
┌──────────────────────────────────────────────────────────┐
│ 1. Sync workspace files to GitHub                        │
│ 2. Update main branch                                    │
│ 3. Create version branch (e.g. branch: v1.2)             │
│ 4. Tag commit & Publish formal GitHub Release            │
└──────────────────────────────────────────────────────────┘
```

---

## 📊 Database Relationships

```
students (1) ──→ (many) responses
                         ↓
                    questions (answered)

students (1) ──→ (many) results

students (1) ──→ (1) student_class_lock

teachers (1) ──→ (many) teacher_assignments
                         ↓
                    classes & subjects

exam_control (singleton) - Global exam state

shuffled_questions - Links students to their shuffled question orders

combined_tests (1) ──→ (many) combined_test_subjects
```

---

## 🔑 Key Features

### 1. Question Shuffling
- Questions appear in random order for each student
- Multiple choice options are also randomized
- Prevents students from using memorized positions

### 2. Negative Marking (Optional)
- Admin can enable/disable
- Wrong answers deduct points
- Formula: `score = correct - (wrong × negative_value)`
- Default: -0.33 per wrong answer

### 3. Reattempt Management
- Students can request reattempt after submission
- Admin can approve/reject requests
- Tracks all reattempt attempts

### 4. Class Lock
- Once a student takes an exam in Class X
- They can only take exams in that class
- Prevents students from switching classes
- Admin must manually unlock if needed

### 5. Combined Tests
- Multi-subject exams
- Questions from multiple subjects in one test
- Admin can set test numbers to map subjects

### 6. Real-time Monitoring
- Admin sees all students taking exam
- Knows who started, who's in progress, who submitted
- Can force-submit a student's exam
- Can view any student's responses

### 7. Exam Time Control
- Admin sets exam duration (default: 60 minutes)
- Countdown timer shows on student exam page
- Auto-submit when time expires
- `auto_submit_expired_exams()` runs automatically

### 8. GitHub Release & Version Control
- Keeps every update (`v1.0`, `v1.1`, `v1.2`, `v1.3`, etc.) separate and downloadable on GitHub
- Dedicated version branches and releases for rollback or reference

---

## 🎨 Frontend Integration

### Key JavaScript Files
- **exam.js** - Handles exam page interactions
  - Timer countdown
  - Question navigation
  - Answer saving
  - Exam submission
  
- **multimodal.js** - Handles multi-subject exam display

### Key HTML Templates
- **exam.html** - Main exam interface
- **results.html** - Results display for students
- **monitoring.html** / **admin_monitoring.html** - Real-time monitoring
- **teacher_monitoring.html** - Teacher's student monitoring

---

## 📁 File Structure

```
RRB_v110/
├── app.py (~5976 lines)
│   ├── Helper functions (password, files, DB)
│   ├── Student routes (login, exam, submit, reattempt)
│   ├── Admin routes (dashboard, questions, monitoring, results, settings)
│   └── Teacher routes (dashboard, monitoring, profiles)
├── push_to_github.py (GitHub versioning, branch & release publisher)
├── Push_Version_To_GitHub.bat (Double-click version push script)
├── RRB_CBT_Manager.bat (Manager launcher script)
├── database.db (SQLite database)
├── static/
│   ├── css/style.css
│   ├── js/exam.js
│   ├── js/multimodal.js
│   ├── student_profile_pic/
│   ├── Teacher_profile_picture/
│   └── uploads/
├── templates/ (HTML templates for student, admin, teacher views)
├── questions/ (CSV question files)
├── backup/ (App backups and historical scripts)
├── exports/ (Generated result exports and reports)
├── uploads/ (Uploaded media and images)
└── requirements.txt (Python dependencies)
```

---

## 🔒 Authentication & Authorization

### Three User Types:
1. **Admin** - `admin_required` decorator
   - Full control over system
   - Exam start/stop
   - Question management
   - Student management
   - Results export

2. **Student** - `student_required` decorator
   - Take exams
   - View results
   - Request reattempt

3. **Teacher** - `teacher_required` decorator
   - View assigned students
   - Monitor student performance
   - View student responses
   - Manage student profiles

### Session Management:
- `session['student_id']` - Student logged in
- `session['admin_logged_in']` - Admin logged in
- `session['teacher_logged_in']` - Teacher logged in

---

## 🛠️ Important Helper Functions

### Image Upload
```python
save_student_picture() → student_profile_pic/student_ID.ext
save_teacher_picture() → Teacher_profile_picture/teacher_ID.ext
save_question_image() → uploads/CLASS_SUBJECT/image.ext
```

### Database
```python
get_db() → Returns connected SQLite database
get_setting(key) → Retrieves app setting
set_setting(key, value) → Saves app setting
```

### Scoring
```python
calculate_score(student_id, subject, negative_marking, negative_value)
→ Computes score and saves to results table
```

### Auto-Submit
```python
auto_submit_expired_exams()
→ Finds exams where time expired and force-submits them
```

---

## 🚀 Version Publishing & Deployment

### To Push a New Version Update (e.g. v1.1, v1.2, v1.3):
1. Double-click **`Push_Version_To_GitHub.bat`** (or run `python push_to_github.py`).
2. Enter the Version Tag (e.g., `v1.2`).
3. Enter description notes (optional).
4. The script will automatically update `main`, create branch `v1.2`, tag the release, and publish a formal GitHub Release on `https://github.com/15btcs0142/RRB_CBT/releases`.

---

## 📞 Troubleshooting

**Problem**: Students can't see questions
- Check: Questions uploaded for that class/subject?
- Check: Exam marked as active?
- Check: Student has correct class/subject?

**Problem**: Exam timer not working
- Check: exam_control.duration set correctly?
- Check: JavaScript errors in console?

**Problem**: Negative marking not applied
- Check: admin_exam_settings enabled negative marking?
- Check: negative_value is set (not 0)?

**Problem**: Students locked to wrong class
- Solution: Delete from student_class_lock table for that student

**Problem**: Shuffled questions not showing
- Check: shuffled_questions table populated?
- Check: Student in current session?

---

## 📝 Version Info
- **App Version**: v1.10
- **Framework**: Flask
- **Database**: SQLite
- **Python**: 3.x
- **GitHub Repository**: https://github.com/15btcs0142/RRB_CBT

