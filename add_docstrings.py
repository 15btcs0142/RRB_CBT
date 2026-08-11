#!/usr/bin/env python3
"""
Script to add descriptive docstrings to all functions in app.py
"""

import re
import os

# Mapping of function names to their descriptions
FUNCTION_DESCRIPTIONS = {
    'hash_password': 'Hash a password using SHA256 encryption.',
    'verify_password': 'Verify if a provided password matches its hashed version.',
    'save_student_picture': 'Save a student profile picture to the file system and return the file path.',
    'save_teacher_picture': 'Save a teacher profile picture to the file system and return the file path.',
    'allowed_image_file': 'Check if a filename has an allowed image extension (png, jpg, jpeg, gif, webp).',
    'save_question_image': 'Save a question image to the appropriate folder based on class and subject.',
    'init_db': 'Initialize the SQLite database with all required tables and schema. Handles migrations from old schemas.',
    'reset_exam_on_startup': 'Reset the exam status on application startup to ensure clean state.',
    'get_db': 'Get a database connection with Row factory enabled for dict-like access.',
    'get_setting': 'Retrieve a setting value from the settings table by key.',
    'set_setting': 'Save or update a setting in the settings table.',
    'admin_required': 'Decorator to enforce admin login requirement on routes.',
    'student_required': 'Decorator to enforce student login requirement on routes.',
    'teacher_required': 'Decorator to enforce teacher login requirement on routes.',
    'index': 'Render the student login page with available classes, subjects, bulletins, and class teachers.',
    'student_login': 'Handle student login and exam registration. Validates class lock, manages reattempt requests.',
    '_get_class_subject_map': 'Helper function to get a mapping of classes to their available subjects.',
    'api_test_numbers': 'API endpoint to retrieve available test numbers for a given class and subject.',
    'generate_combined_test': 'API endpoint to generate a combined test with multiple subjects.',
    'api_check_test_no': 'API endpoint to check if a test number is valid for the given class and subject.',
    'waiting': 'Render the waiting page for students to wait for exam to start.',
    'check_exam_status': 'API endpoint to check if the exam has started.',
    'exam': 'Render the exam page with shuffled questions for the student.',
    'get_questions': 'API endpoint to retrieve exam questions with shuffled options for display.',
    'save_answer': 'API endpoint to save a student\'s answer to a question.',
    'submit_exam': 'API endpoint to submit the exam and calculate results.',
    'submitted': 'Render the page shown after exam submission with reattempt option.',
    'request_reattempt': 'API endpoint for students to request reattempt of an exam.',
    'check_reattempt_status': 'API endpoint to check the status of a reattempt request.',
    'get_exam_time': 'API endpoint to get remaining exam time and status.',
    'admin_login': 'Handle admin login and render admin login page.',
    'admin_dashboard': 'Render the admin dashboard with statistics.',
    'admin_logout': 'Handle admin logout.',
    'exam_status': 'Render the exam status page showing all students and their progress.',
    'start_exam': 'API endpoint for admin to start the exam.',
    'admin_exam_settings': 'Render page for admin to configure exam settings (duration, negative marking, etc).',
    'admin_force_submit_student': 'API endpoint for admin to force submit exam for a specific student.',
    'stop_exam': 'API endpoint for admin to stop the exam.',
    'admin_settings': 'Render and handle admin settings page (school name, logo, etc).',
    'generate_question_paper': 'Render page to generate a PDF of the question paper.',
    'questions': 'Render the admin questions management page.',
    'questions_data': 'API endpoint to fetch questions data with filtering and pagination.',
    'add_question': 'API endpoint to add a new question to the database.',
    'update_question': 'API endpoint to update an existing question.',
    'delete_question': 'API endpoint to delete a question by ID.',
    'delete_questions_by_class_subject': 'API endpoint to delete all questions for a specific class and subject.',
    'upload_csv': 'API endpoint to upload questions from a CSV file.',
    'manage_students': 'Render the admin students management page.',
    'students_data': 'API endpoint to fetch students data with filtering and pagination.',
    'delete_student': 'API endpoint to delete a student record.',
    'monitoring': 'Render the exam monitoring page showing real-time student status.',
    'monitoring_data': 'API endpoint to fetch real-time monitoring data.',
    'evaluate': 'Render the page for evaluating descriptive responses.',
    'results_page': 'Render the results page showing all test results.',
    'results_data': 'API endpoint to fetch results data with filtering and pagination.',
    'export_results_page': 'Render the page for exporting results.',
    'export_results': 'API endpoint to export results to Excel format.',
    'view_student_responses': 'Render page to view all responses of a specific student.',
    'teacher_login': 'Handle teacher login.',
    'teacher_dashboard': 'Render the teacher dashboard.',
    'teacher_logout': 'Handle teacher logout.',
    'teacher_monitoring': 'Render the teacher monitoring page for their assigned students.',
    'teacher_students': 'Render the page showing teacher\'s assigned students.',
    'teacher_student_profile': 'Render the profile page of a student (teacher view).',
    'teacher_test_history': 'Render the page showing test history for teacher\'s classes.',
    'teacher_test_responses': 'Render page to view student responses to questions.',
    'auto_submit_expired_exams': 'Auto-submit exams that have exceeded the time limit.',
    'calculate_score': 'Calculate the score and percentage for a student\'s exam.',
    'shuffle_questions_for_student': 'Create a shuffled version of questions for a student.',
    'get_shuffled_questions': 'Retrieve the shuffled questions for a student.',
    'allowed_file': 'Check if a filename has an allowed extension (CSV).',
    'create_test_papers_list': 'Create a mapping of available test papers by class and subject.',
    'get_class_subject_map': 'Get the class-subject mapping from the database.',
}

def add_docstrings_to_file(file_path):
    """
    Read the file and add docstrings to functions that don't have them.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    new_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Check if this is a function definition (considering decorators too)
        if re.match(r'^def\s+\w+\s*\(', line):
            # Extract function name
            func_match = re.match(r'^def\s+(\w+)\s*\(', line)
            if func_match:
                func_name = func_match.group(1)
                
                # Add the function definition line
                new_lines.append(line)
                i += 1
                
                # Check if the next non-empty line is a docstring
                peek_idx = i
                while peek_idx < len(lines) and lines[peek_idx].strip() == '':
                    new_lines.append(lines[peek_idx])
                    i += 1
                    peek_idx += 1
                
                # Check if there's already a docstring
                has_docstring = False
                if peek_idx < len(lines):
                    next_line = lines[peek_idx].strip()
                    if next_line.startswith('"""') or next_line.startswith("'''"):
                        has_docstring = True
                
                # If no docstring, add one
                if not has_docstring and func_name in FUNCTION_DESCRIPTIONS:
                    indent = len(line) - len(line.lstrip())
                    indent_str = ' ' * (indent + 4)
                    description = FUNCTION_DESCRIPTIONS[func_name]
                    
                    new_lines.append(f'{indent_str}"""\n')
                    new_lines.append(f'{indent_str}{description}\n')
                    new_lines.append(f'{indent_str}"""\n')
                
                continue
        
        new_lines.append(line)
        i += 1
    
    return ''.join(new_lines)

def main():
    app_file = 'app.py'
    backup_file = 'app_backup.py'
    
    # Create backup
    if os.path.exists(app_file):
        with open(app_file, 'r', encoding='utf-8') as f:
            backup_content = f.read()
        with open(backup_file, 'w', encoding='utf-8') as f:
            f.write(backup_content)
        print(f"✓ Backup created: {backup_file}")
    
    # Add docstrings
    new_content = add_docstrings_to_file(app_file)
    
    # Write back to the file
    with open(app_file, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"✓ Successfully added docstrings to {app_file}")
    print(f"✓ Functions documented: {len([k for k in FUNCTION_DESCRIPTIONS.keys()])}")

if __name__ == '__main__':
    main()
