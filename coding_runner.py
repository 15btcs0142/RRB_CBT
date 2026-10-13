"""
coding_runner.py - Lightweight, containerless Code Execution Runner for RRB CBT Platform.
Executes Python, Java, C, and C++ student submissions safely in separate subprocesses.
"""

import os
import sys
import uuid
import time
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor

# Maximum workers for concurrent student code executions
MAX_WORKERS = 4
EXECUTOR = ThreadPoolExecutor(max_workers=MAX_WORKERS)

# Max allowed output size in bytes (100 KB limit to prevent infinite print loops)
MAX_OUTPUT_BYTES = 100 * 1024

def normalize_output(text: str) -> str:
    """Normalize output for comparison: unify newlines, rstrip each line, and strip outer spaces."""
    if not text:
        return ""
    lines = text.replace('\r\n', '\n').replace('\r', '\n').strip().split('\n')
    return '\n'.join(line.rstrip() for line in lines)

def check_compiler_availability(language: str) -> tuple[bool, str]:
    """Check if the compiler/interpreter for the requested language is available in PATH."""
    lang = language.lower().strip()
    if lang == 'python':
        return True, sys.executable
    elif lang == 'java':
        javac_path = shutil.which('javac')
        java_path = shutil.which('java')
        if not javac_path or not java_path:
            return False, "Java Development Kit (JDK) not found in PATH. Please install JDK (javac & java)."
        return True, javac_path
    elif lang == 'c':
        gcc_path = shutil.which('gcc') or shutil.which('clang') or shutil.which('cl')
        if not gcc_path:
            return False, "C compiler (gcc/clang/cl) not found in PATH. Please install MinGW-w64 or GCC and add to PATH."
        return True, gcc_path
    elif lang in ('cpp', 'c++'):
        gpp_path = shutil.which('g++') or shutil.which('clang++') or shutil.which('cl')
        if not gpp_path:
            return False, "C++ compiler (g++/clang++/cl) not found in PATH. Please install MinGW-w64 or G++ and add to PATH."
        return True, gpp_path
    else:
        return False, f"Unsupported programming language: '{language}'"

def run_code(code: str, language: str, stdin_input: str = '', timeout_seconds: int = 5) -> dict:
    """
    Compile and execute code safely in a per-submission temporary directory.
    
    Parameters:
        code (str): Source code string
        language (str): 'python', 'java', 'c', or 'cpp'
        stdin_input (str): Standard input provided to the program
        timeout_seconds (int): Maximum execution time allowed before killing the process
        
    Returns:
        dict: { status, stdout, stderr, execution_time_ms, error_message }
    """
    lang = language.lower().strip()
    
    # Verify compiler/interpreter availability
    available, msg = check_compiler_availability(lang)
    if not available:
        return {
            'status': 'Compiler Not Found',
            'stdout': '',
            'stderr': msg,
            'execution_time_ms': 0.0,
            'error_message': msg
        }
        
    # Create unique isolation temporary directory
    base_temp = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'scratch', 'temp_code_exec')
    os.makedirs(base_temp, exist_ok=True)
    temp_dir = os.path.join(base_temp, f"exec_{uuid.uuid4().hex}")
    os.makedirs(temp_dir, exist_ok=True)
    
    start_time = time.time()
    try:
        if lang == 'python':
            src_file = os.path.join(temp_dir, 'solution.py')
            with open(src_file, 'w', encoding='utf-8') as f:
                f.write(code)
            cmd = [sys.executable, src_file]
            return _execute_subprocess(cmd, stdin_input, temp_dir, timeout_seconds, start_time)
            
        elif lang == 'java':
            # Extract public class name if available, default to Main
            class_name = 'Main'
            for line in code.splitlines():
                line_str = line.strip()
                if line_str.startswith('public class '):
                    parts = line_str.split()
                    if len(parts) >= 3:
                        class_name = parts[2].split('{')[0].strip()
                        break
            
            src_file = os.path.join(temp_dir, f"{class_name}.java")
            with open(src_file, 'w', encoding='utf-8') as f:
                f.write(code)
                
            # Compile step
            compile_cmd = ['javac', src_file]
            comp_res = subprocess.run(
                compile_cmd, cwd=temp_dir, capture_output=True, text=True, timeout=15
            )
            if comp_res.returncode != 0:
                return {
                    'status': 'Compile Error',
                    'stdout': comp_res.stdout[:MAX_OUTPUT_BYTES],
                    'stderr': comp_res.stderr[:MAX_OUTPUT_BYTES],
                    'execution_time_ms': (time.time() - start_time) * 1000,
                    'error_message': 'Java Compilation Failed'
                }
                
            # Run step
            run_cmd = ['java', '-cp', temp_dir, class_name]
            return _execute_subprocess(run_cmd, stdin_input, temp_dir, timeout_seconds, start_time)
            
        elif lang == 'c':
            src_file = os.path.join(temp_dir, 'solution.c')
            exe_file = os.path.join(temp_dir, 'solution.exe' if os.name == 'nt' else 'solution')
            with open(src_file, 'w', encoding='utf-8') as f:
                f.write(code)
                
            compile_cmd = ['gcc', '-O2', src_file, '-o', exe_file]
            comp_res = subprocess.run(
                compile_cmd, cwd=temp_dir, capture_output=True, text=True, timeout=15
            )
            if comp_res.returncode != 0:
                return {
                    'status': 'Compile Error',
                    'stdout': comp_res.stdout[:MAX_OUTPUT_BYTES],
                    'stderr': comp_res.stderr[:MAX_OUTPUT_BYTES],
                    'execution_time_ms': (time.time() - start_time) * 1000,
                    'error_message': 'C Compilation Failed'
                }
                
            return _execute_subprocess([exe_file], stdin_input, temp_dir, timeout_seconds, start_time)
            
        elif lang in ('cpp', 'c++'):
            src_file = os.path.join(temp_dir, 'solution.cpp')
            exe_file = os.path.join(temp_dir, 'solution.exe' if os.name == 'nt' else 'solution')
            with open(src_file, 'w', encoding='utf-8') as f:
                f.write(code)
                
            compile_cmd = ['g++', '-O2', src_file, '-o', exe_file]
            comp_res = subprocess.run(
                compile_cmd, cwd=temp_dir, capture_output=True, text=True, timeout=15
            )
            if comp_res.returncode != 0:
                return {
                    'status': 'Compile Error',
                    'stdout': comp_res.stdout[:MAX_OUTPUT_BYTES],
                    'stderr': comp_res.stderr[:MAX_OUTPUT_BYTES],
                    'execution_time_ms': (time.time() - start_time) * 1000,
                    'error_message': 'C++ Compilation Failed'
                }
                
            return _execute_subprocess([exe_file], stdin_input, temp_dir, timeout_seconds, start_time)
            
    finally:
        # Guarantee cleanup of temporary submission folder
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

def _execute_subprocess(cmd: list, stdin_input: str, cwd: str, timeout_seconds: int, start_time: float) -> dict:
    """Helper to run a subprocess with stdin, timeout, and output capping."""
    try:
        proc = subprocess.Popen(
            cmd,
            cwd=cwd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        stdout_data, stderr_data = proc.communicate(input=stdin_input, timeout=timeout_seconds)
        exec_time_ms = (time.time() - start_time) * 1000
        
        # Enforce output cap
        stdout_str = stdout_data[:MAX_OUTPUT_BYTES] if stdout_data else ""
        stderr_str = stderr_data[:MAX_OUTPUT_BYTES] if stderr_data else ""
        
        if proc.returncode != 0:
            return {
                'status': 'Runtime Error',
                'stdout': stdout_str,
                'stderr': stderr_str or f"Process exited with non-zero exit code {proc.returncode}",
                'execution_time_ms': exec_time_ms,
                'error_message': f"Runtime Exception (Exit code {proc.returncode})"
            }
            
        return {
            'status': 'Success',
            'stdout': stdout_str,
            'stderr': stderr_str,
            'execution_time_ms': exec_time_ms,
            'error_message': ''
        }
        
    except subprocess.TimeoutExpired:
        try:
            proc.kill()
            proc.communicate()
        except Exception:
            pass
        return {
            'status': 'Time Limit Exceeded',
            'stdout': '',
            'stderr': f'Time Limit Exceeded ({timeout_seconds} seconds limit)',
            'execution_time_ms': timeout_seconds * 1000.0,
            'error_message': f'Execution timed out after {timeout_seconds} seconds'
        }
    except Exception as e:
        return {
            'status': 'Execution Error',
            'stdout': '',
            'stderr': str(e),
            'execution_time_ms': (time.time() - start_time) * 1000,
            'error_message': f"Subprocess exception: {str(e)}"
        }

def evaluate_submission(code: str, language: str, test_cases: list, question_marks: float = 10.0, timeout_seconds: int = 5) -> dict:
    """
    Evaluates student code against a list of test cases (dict with 'input', 'expected_output', 'marks_weight', 'is_hidden').
    
    Returns:
        dict: { status, passed_count, total_test_cases, marks_awarded, results: [...] }
    """
    if not test_cases:
        # If no test cases provided, run once with empty input
        res = run_code(code, language, stdin_input='', timeout_seconds=timeout_seconds)
        return {
            'status': res['status'],
            'passed_count': 1 if res['status'] == 'Success' else 0,
            'total_test_cases': 1,
            'marks_awarded': question_marks if res['status'] == 'Success' else 0.0,
            'results': [res]
        }
        
    total_weight = sum(float(tc.get('marks_weight', 1.0)) for tc in test_cases) or 1.0
    passed_weight = 0.0
    passed_count = 0
    case_results = []
    overall_status = 'Passed'
    
    for idx, tc in enumerate(test_cases):
        inp = tc.get('input', '')
        exp = tc.get('expected_output', '')
        weight = float(tc.get('marks_weight', 1.0))
        is_hidden = bool(tc.get('is_hidden', 0))
        
        run_res = run_code(code, language, stdin_input=inp, timeout_seconds=timeout_seconds)
        
        actual_norm = normalize_output(run_res['stdout'])
        expected_norm = normalize_output(exp)
        
        passed = (run_res['status'] == 'Success') and (actual_norm == expected_norm)
        if passed:
            passed_count += 1
            passed_weight += weight
        else:
            if run_res['status'] in ('Compile Error', 'Compiler Not Found'):
                overall_status = run_res['status']
            elif overall_status == 'Passed':
                overall_status = 'Failed' if run_res['status'] == 'Success' else run_res['status']
                
        case_results.append({
            'test_case_id': tc.get('id', idx + 1),
            'is_hidden': is_hidden,
            'passed': passed,
            'status': run_res['status'] if run_res['status'] != 'Success' else ('Passed' if passed else 'Wrong Answer'),
            'input': inp if not is_hidden else '[Hidden]',
            'expected_output': exp if not is_hidden else '[Hidden]',
            'actual_output': run_res['stdout'] if not is_hidden else '[Hidden]',
            'execution_time_ms': run_res['execution_time_ms'],
            'stderr': run_res['stderr'] if not is_hidden else ''
        })
        
    marks_awarded = round((passed_weight / total_weight) * question_marks, 2)
    if passed_count == len(test_cases):
        overall_status = 'Passed'
    elif passed_count > 0 and overall_status not in ('Compile Error', 'Compiler Not Found'):
        overall_status = 'Partially Passed'
        
    return {
        'status': overall_status,
        'passed_count': passed_count,
        'total_test_cases': len(test_cases),
        'marks_awarded': marks_awarded,
        'results': case_results
    }
