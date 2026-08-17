# RRB CBT v1.13 - Enhanced Edition
## School Online Examination System

---

## 🆕 What's New in v1.13

### 1. Copy/Paste Prevention in Exam
- Right-click disabled
- Ctrl+C, Ctrl+V, Ctrl+X, Ctrl+A all blocked
- Text selection disabled
- Drag-and-drop blocked

### 2. Teacher Portal (`/teacher`)
- Login via **mobile number + password**
- View assigned classes and subjects
- Dashboard with quick access to all features

### 3. AI Test Generation
- Teachers can create tests using **Claude AI**
- Specify: Class, Subject, Chapter, No. of Questions, and instructions
- Also supports manual **CSV upload**

### 4. Student Profiles
- Full details: Name, Admission No, Class, Section, DOB, House, Parents Name, Address
- Profile picture upload
- Class teachers can **edit all details**
- Subject teachers can **view only their subject progress**

### 5. Test History & Progress Charts
- Every test result saved with date
- Interactive **line chart** showing percentage over time per subject
- Click any test to view full question paper + responses
- PDF report generation per student

### 6. Teacher Management (Admin)
- Add teachers individually (fill form)
- **Bulk import via CSV** (columns: name, mobile, password, email)
- Assign class + subject per teacher
- Mark as **Class Teacher** (sees all subjects) or Subject Teacher

---

## 🚀 How to Run

### Windows
Double-click `RRB_CBT_Manager.bat`

### Manual
```bash
pip install -r requirements.txt
python app.py
```
Then open: **http://localhost:5000**

---

## 🔐 Login URLs
| Role    | URL          |
|---------|--------------|
| Student | `/`          |
| Admin   | `/admin`     |
| Teacher | `/teacher`   |

**Default admin password:** Set in your existing `app.py` (unchanged)

---

## 📁 CSV Format for Bulk Teacher Import
```
name,mobile,password,email
Amit Sharma,9876543210,pass123,amit@school.com
Priya Singh,9123456789,priya456,
```
Password column is optional — defaults to `teacher123`

---

## 📊 Student Profile CSV (Questions)
```
question,option a,option b,option c,option d,correct_answer
What is 2+2?,3,4,5,6,4
```

---

## ⚙️ Setting up Anthropic API (for AI Test Generation)
1. Get your API key from https://console.anthropic.com
2. Set the environment variable:
   - **Windows:** `set ANTHROPIC_API_KEY=your-key-here`
   - **Linux/Mac:** `export ANTHROPIC_API_KEY=your-key-here`
3. Or add to `RRB_CBT_Manager.bat`:
   ```
   set ANTHROPIC_API_KEY=your-key-here
   ```

---

## 📂 Project Structure
```
RRB_cbt-v1.13/
├── app.py                    ← Main application
├── database.db               ← SQLite database
├── requirements.txt
├── templates/
│   ├── teacher_login.html
│   ├── teacher_dashboard.html
│   ├── teacher_create_test.html
│   ├── teacher_students.html
│   ├── teacher_student_profile.html
│   ├── teacher_test_responses.html
│   ├── admin_teachers.html   ← NEW
│   └── ... (existing templates)
├── static/
│   ├── css/style.css
│   ├── js/exam.js            ← Updated with copy prevention
│   └── uploads/
│       └── students/         ← Student profile pictures
└── questions/                ← CSV question files
```
