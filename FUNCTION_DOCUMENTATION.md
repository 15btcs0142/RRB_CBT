# RRB CBT Application - Function Documentation Guide

## Overview
This document provides a comprehensive list of all functions in `app.py` with their descriptions to help you understand the application structure.

---

## 📋 Helper Functions (Utilities)

### Password & Security
- **`hash_password(password)`** - Hash a password using SHA256 encryption.
- **`verify_password(password, hashed)`** - Verify if a provided password matches its hashed version.

### File Handling
- **`save_student_picture(file, student_id)`** - Save a student profile picture to the file system and return the file path.
- **`save_teacher_picture(file, teacher_id)`** - Save a teacher profile picture to the file system and return the file path.
- **`allowed_image_file(filename)`** - Check if a filename has an allowed image extension (png, jpg, jpeg, gif, webp).
- **`save_question_image(file, class_, subject, question_id)`** - Save a question image to the appropriate folder based on class and subject.
- **`allowed_file(filename)`** - Check if a filename has an allowed extension (CSV).

### Database Operations
- **`init_db()`** - Initialize the SQLite database with all required tables and schema. Handles migrations from old schemas.
- **`reset_exam_on_startup()`** - Reset the exam status on application startup to ensure clean state.
- **`get_db()`** - Get a database connection with Row factory enabled for dict-like access.
- **`get_setting(key, default='')`** - Retrieve a setting value from the settings table by key.
- **`set_setting(key, value)`** - Save or update a setting in the settings table.

### Authentication Decorators
- **`admin_required(f)`** - Decorator to enforce admin login requirement on routes.
- **`student_required(f)`** - Decorator to enforce student login requirement on routes.
- **`teacher_required(f)`** - Decorator to enforce teacher login requirement on routes.

### Helper Utilities
- **`_get_class_subject_map()`** - Helper function to get a mapping of classes to their available subjects.
- **`get_class_subject_map()`** - Get the class-subject mapping from the database.
- **`create_test_papers_list()`** - Create a mapping of available test papers by class and subject.
- **`shuffle_questions_for_student(student_id, question_ids)`** - Create a shuffled version of questions for a student.
- **`get_shuffled_questions(student_id, question_ids)`** - Retrieve the shuffled questions for a student.
- **`calculate_score(student_id, subject, negative_marking, negative_value)`** - Calculate the score and percentage for a student's exam.
- **`auto_submit_expired_exams()`** - Auto-submit exams that have exceeded the time limit.

---

## 🎓 Student Routes & Pages

### Student Login & Registration
- **`index()`** [GET] - Render the student login page with available classes, subjects, bulletins, and class teachers.
- **`student_login()`** [POST] - Handle student login and exam registration. Validates class lock, manages reattempt requests.

### Exam Workflow
- **`api_test_numbers()`** [GET] - API endpoint to retrieve available test numbers for a given class and subject.
- **`api_check_test_no()`** [GET] - API endpoint to check if a test number is valid for the given class and subject.
- **`waiting()`** [GET] - Render the waiting page for students to wait for exam to start.
- **`check_exam_status()`** [GET] - API endpoint to check if the exam has started.
- **`exam()`** [GET] - Render the exam page with shuffled questions for the student.
- **`get_questions()`** [GET] - API endpoint to retrieve exam questions with shuffled options for display.
- **`get_exam_time()`** [GET] - API endpoint to get remaining exam time and status.
- **`save_answer()`** [POST] - API endpoint to save a student's answer to a question.
- **`submit_exam()`** [POST] - API endpoint to submit the exam and calculate results.

### After Exam
- **`submitted()`** [GET] - Render the page shown after exam submission with reattempt option.
- **`request_reattempt()`** [POST] - API endpoint for students to request reattempt of an exam.
- **`check_reattempt_status()`** [GET] - API endpoint to check the status of a reattempt request.

### Combined Tests (Multi-Subject)
- **`generate_combined_test()`** [POST] - API endpoint to generate a combined test with multiple subjects.

---

## 🔐 Admin Routes & Pages

### Admin Authentication
- **`admin_login()`** [GET/POST] - Handle admin login and render admin login page.
- **`admin_logout()`** [GET] - Handle admin logout.

### Admin Dashboard & Monitoring
- **`admin_dashboard()`** [GET] - Render the admin dashboard with statistics.
- **`exam_status()`** [GET] - Render the exam status page showing all students and their progress.
- **`monitoring()`** [GET] - Render the exam monitoring page showing real-time student status.
- **`monitoring_data()`** [GET] - API endpoint to fetch real-time monitoring data.

### Exam Control
- **`start_exam()`** [POST] - API endpoint for admin to start the exam.
- **`stop_exam()`** [GET] - API endpoint for admin to stop the exam.
- **`admin_exam_settings()`** [GET] - Render page for admin to configure exam settings (duration, negative marking, etc).
- **`admin_force_submit_student(student_id)`** [POST] - API endpoint for admin to force submit exam for a specific student.

### Settings Management
- **`admin_settings()`** [GET/POST] - Render and handle admin settings page (school name, logo, etc).

### Questions Management
- **`questions()`** [GET] - Render the admin questions management page.
- **`questions_data()`** [GET] - API endpoint to fetch questions data with filtering and pagination.
- **`add_question()`** [POST] - API endpoint to add a new question to the database.
- **`update_question(qid)`** [POST] - API endpoint to update an existing question.
- **`delete_question(qid)`** [DELETE] - API endpoint to delete a question by ID.
- **`delete_questions_by_class_subject()`** [POST] - API endpoint to delete all questions for a specific class and subject.
- **`upload_csv()`** [POST] - API endpoint to upload questions from a CSV file.
- **`generate_question_paper()`** [GET] - Render page to generate a PDF of the question paper.

### Students Management
- **`manage_students()`** [GET] - Render the admin students management page.
- **`students_data()`** [GET] - API endpoint to fetch students data with filtering and pagination.
- **`delete_student(student_id)`** [DELETE] - API endpoint to delete a student record.

### Results & Evaluation
- **`evaluate()`** [GET] - Render the page for evaluating descriptive responses.
- **`results_page()`** [GET] - Render the results page showing all test results.
- **`results_data()`** [GET] - API endpoint to fetch results data with filtering and pagination.
- **`view_student_responses(student_id)`** [GET] - Render page to view all responses of a specific student.

### Results Export
- **`export_results_page()`** [GET] - Render the page for exporting results.
- **`export_results()`** [GET] - API endpoint to export results to Excel format.

---

## 👨‍🏫 Teacher Routes & Pages

### Teacher Authentication
- **`teacher_login()`** [GET/POST] - Handle teacher login.
- **`teacher_logout()`** [GET] - Handle teacher logout.

### Teacher Dashboard & Monitoring
- **`teacher_dashboard()`** [GET] - Render the teacher dashboard.
- **`teacher_monitoring()`** [GET] - Render the teacher monitoring page for their assigned students.

### Students & Classes
- **`teacher_students()`** [GET] - Render the page showing teacher's assigned students.
- **`teacher_student_profile(student_id)`** [GET] - Render the profile page of a student (teacher view).

### Test History & Responses
- **`teacher_test_history()`** [GET] - Render the page showing test history for teacher's classes.
- **`teacher_test_responses()`** [GET] - Render page to view student responses to questions.
- **`teacher_profile()`** [GET/POST] - Render and handle teacher profile page.

---

## 🗄️ Database Tables & Structure

### Main Tables
1. **students** - Student information and exam status
2. **questions** - Question bank with options and correct answers
3. **responses** - Student answers to questions
4. **results** - Exam results and scores
5. **exam_control** - Global exam settings (active/inactive, duration, etc)
6. **settings** - Application settings (school name, logo, etc)
7. **shuffled_questions** - Question shuffling order per student
8. **teachers** - Teacher information
9. **teacher_assignments** - Teacher-class-subject assignments
10. **test_papers** - Test paper records
11. **bulletins** - Announcements/Bulletins
12. **reattempt_requests** - Student reattempt requests
13. **combined_tests** - Multi-subject tests
14. **combined_test_subjects** - Subjects in combined tests
15. **test_generation_history** - AI-generated test history
16. **student_class_lock** - Class lock for students

---

## 🚀 Application Flow

### Student Exam Flow
1. **Student lands on home page** → `index()`
2. **Student logs in** → `student_login()`
3. **Checks exam status** → `check_exam_status()`
4. **Waits for exam to start** → `waiting()`
5. **Enters exam** → `exam()`
6. **Gets questions** → `get_questions()`
7. **Saves answers** → `save_answer()` (multiple times)
8. **Submits exam** → `submit_exam()`
9. **Views result** → `submitted()`
10. **Can request reattempt** → `request_reattempt()`

### Admin Control Flow
1. **Admin logs in** → `admin_login()`
2. **Views dashboard** → `admin_dashboard()`
3. **Manages questions** → `questions()`, `add_question()`, `upload_csv()`
4. **Monitors students** → `monitoring()`, `monitoring_data()`
5. **Starts/stops exam** → `start_exam()`, `stop_exam()`
6. **Views results** → `results_page()`, `export_results()`
7. **Logs out** → `admin_logout()`

### Teacher Control Flow
1. **Teacher logs in** → `teacher_login()`
2. **Views dashboard** → `teacher_dashboard()`
3. **Views students** → `teacher_students()`
4. **Monitors their students** → `teacher_monitoring()`
5. **Checks responses** → `teacher_test_responses()`

---

## 📝 Notes

- **Backup File**: A backup of your original app.py has been created as `app_backup.py`
- **Documentation**: Each function now has a docstring describing its purpose
- **Total Functions Documented**: 73 functions
- **Decorators**: Routes use `@student_required`, `@admin_required`, and `@teacher_required` decorators for authentication
- **API Endpoints**: Many routes are JSON API endpoints (prefixed with `/api/`)

---

## 🔍 How to Search

To find a specific function:
1. Use Ctrl+F (Cmd+F on Mac) in your editor
2. Search for `def function_name()` to find the definition
3. The docstring will appear right below the function signature

Example:
```python
def save_answer():
    """
    API endpoint to save a student's answer to a question.
    """
    # Function implementation...
```

---

**Generated**: $(date) | **Version**: 1.0
