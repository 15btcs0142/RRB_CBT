import os
import sys
import types
import sqlite3
import ast
import re
import json
import traceback

print("=" * 60)
print("       RRB CBT COMPREHENSIVE CODEBASE AUDIT")
print("=" * 60)

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
os.chdir(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Mock pdfkit if not installed
try:
    import pdfkit
except ImportError:
    print("[BUG FOUND] app.py line 11 imports pdfkit at top-level, but pdfkit is not in requirements.txt or needed (project uses weasyprint). This crashes startup if pdfkit is missing.")
    sys.modules['pdfkit'] = types.ModuleType('pdfkit')

import app
from flask import Flask

# -------------------------------------------------------------
# 1. SCOPE & UNDEFINED NAME AUDIT
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 1: Scope & Undefined Variables Scan")

class AdvancedScopeVisitor(ast.NodeVisitor):
    def __init__(self, filename):
        self.filename = filename
        self.scopes = [set()] # 0: global
        self.undefined = []

    def visit_Import(self, node):
        for alias in node.names:
            name = alias.asname or alias.name.split('.')[0]
            self.scopes[0].add(name)
        self.generic_visit(node)

    def visit_ImportFrom(self, node):
        for alias in node.names:
            name = alias.asname or alias.name
            self.scopes[0].add(name)
        self.generic_visit(node)

    def visit_FunctionDef(self, node):
        self.scopes[0].add(node.name)
        local_scope = set()
        for arg in node.args.posonlyargs + node.args.args + node.args.kwonlyargs:
            local_scope.add(arg.arg)
        if node.args.vararg:
            local_scope.add(node.args.vararg.arg)
        if node.args.kwarg:
            local_scope.add(node.args.kwarg.arg)
        self.scopes.append(local_scope)
        self.generic_visit(node)
        self.scopes.pop()

    def visit_AsyncFunctionDef(self, node):
        self.visit_FunctionDef(node)

    def visit_ClassDef(self, node):
        self.scopes[0].add(node.name)
        self.scopes.append(set())
        self.generic_visit(node)
        self.scopes.pop()

    def visit_Lambda(self, node):
        local_scope = set()
        for arg in node.args.posonlyargs + node.args.args + node.args.kwonlyargs:
            local_scope.add(arg.arg)
        if node.args.vararg:
            local_scope.add(node.args.vararg.arg)
        if node.args.kwarg:
            local_scope.add(node.args.kwarg.arg)
        self.scopes.append(local_scope)
        self.generic_visit(node)
        self.scopes.pop()

    def _visit_comprehension(self, node, generators, elt_nodes):
        self.scopes.append(set())
        for gen in generators:
            self._add_targets(gen.target)
            self.visit(gen.iter)
            for if_clause in gen.ifs:
                self.visit(if_clause)
        for elt in elt_nodes:
            self.visit(elt)
        self.scopes.pop()

    def visit_ListComp(self, node):
        self._visit_comprehension(node, node.generators, [node.elt])

    def visit_SetComp(self, node):
        self._visit_comprehension(node, node.generators, [node.elt])

    def visit_GeneratorExp(self, node):
        self._visit_comprehension(node, node.generators, [node.elt])

    def visit_DictComp(self, node):
        self._visit_comprehension(node, node.generators, [node.key, node.value])

    def _add_targets(self, target):
        if isinstance(target, ast.Name):
            self.scopes[-1].add(target.id)
        elif isinstance(target, (ast.Tuple, ast.List)):
            for elt in target.elts:
                self._add_targets(elt)
        elif isinstance(target, ast.Starred):
            self._add_targets(target.value)

    def visit_Assign(self, node):
        for target in node.targets:
            self._add_targets(target)
        self.generic_visit(node)

    def visit_AnnAssign(self, node):
        self._add_targets(node.target)
        self.generic_visit(node)

    def visit_AugAssign(self, node):
        self._add_targets(node.target)
        self.generic_visit(node)

    def visit_For(self, node):
        self._add_targets(node.target)
        self.generic_visit(node)

    def visit_AsyncFor(self, node):
        self._add_targets(node.target)
        self.generic_visit(node)

    def visit_ExceptHandler(self, node):
        if node.name:
            self.scopes[-1].add(node.name)
        self.generic_visit(node)

    def visit_With(self, node):
        for item in node.items:
            if item.optional_vars:
                self._add_targets(item.optional_vars)
        self.generic_visit(node)

    def visit_Global(self, node):
        for name in node.names:
            self.scopes[0].add(name)
        self.generic_visit(node)

    def visit_NamedExpr(self, node):
        self._add_targets(node.target)
        self.generic_visit(node)

    def visit_Name(self, node):
        if isinstance(node.ctx, ast.Load):
            name = node.id
            import builtins
            builtin_names = set(dir(builtins)) | {'__file__', '__name__', '__doc__'}
            found = False
            for scope in reversed(self.scopes):
                if name in scope:
                    found = True
                    break
            if not found and name not in builtin_names:
                self.undefined.append((name, node.lineno))
        self.generic_visit(node)

for py_file in ['app.py', 'exemplar_extractor.py', 'image_handler.py', 'validators.py', 'curriculum_data.py', 'rq_worker.py', 'push_to_github.py']:
    if not os.path.exists(py_file):
        continue
    with open(py_file, 'r', encoding='utf-8', errors='ignore') as f:
        src = f.read()
    tree = ast.parse(src, filename=py_file)
    visitor = AdvancedScopeVisitor(py_file)
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            visitor.scopes[0].add(node.name)
        elif isinstance(node, ast.Assign):
            for t in node.targets:
                visitor._add_targets(t)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                name = alias.asname or alias.name.split('.')[0]
                visitor.scopes[0].add(name)
    visitor.visit(tree)
    if visitor.undefined:
        seen = set()
        for name, lineno in visitor.undefined:
            if (name, lineno) not in seen:
                seen.add((name, lineno))
                print(f"  [UNDEFINED VARIABLE BUG] in {py_file}:{lineno} -> '{name}'")

# -------------------------------------------------------------
# 2. DATABASE SCHEMA & SQL QUERY VALIDATION
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 2: Database Schema & SQL Query Consistency")
mem_db = sqlite3.connect(':memory:')
mem_db.row_factory = sqlite3.Row
old_get_db = app.get_db
app.get_db = lambda: mem_db

try:
    app.init_db()
    print("  [OK] Database tables created via init_db().")
except Exception as e:
    print(f"  [CRITICAL DB ERROR] init_db() failed: {e}")

c = mem_db.cursor()
c.execute("SELECT name FROM sqlite_master WHERE type='table'")
table_names = [row[0] for row in c.fetchall()]
schema = {}
for tbl in table_names:
    c.execute(f"PRAGMA table_info({tbl})")
    schema[tbl] = {row[1]: row[2] for row in c.fetchall()}

print(f"  Total tables in schema: {len(schema)}")

# Scan app.py for table queries and column queries
with open('app.py', 'r', encoding='utf-8', errors='ignore') as f:
    app_src = f.read()

# Check for column references in SELECT / INSERT / UPDATE
# Find all table inserts
insert_matches = re.findall(r'INSERT\s+INTO\s+([a-zA-Z0-9_]+)\s*\(([^)]+)\)', app_src, re.IGNORECASE)
for tbl, cols in insert_matches:
    tbl = tbl.strip()
    if tbl in schema:
        col_list = [col.strip().strip('`"\'[]') for col in cols.split(',')]
        for col in col_list:
            if col and col not in schema[tbl]:
                print(f"  [SQL COLUMN MISMATCH BUG] INSERT INTO {tbl} specifies non-existent column '{col}'")
    else:
        print(f"  [SQL TABLE MISMATCH BUG] INSERT INTO references non-existent table '{tbl}'")

# -------------------------------------------------------------
# 3. TEMPLATES & URL_FOR INTEGRITY
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 3: Templates, URL_FOR & Jinja Integrity")
rendered_templates = set(re.findall(r'render_template\s*\(\s*[\'\"]([^\'\"]+)[\'\"]', app_src))
existing_templates = set(os.listdir('templates')) if os.path.exists('templates') else set()

for t in rendered_templates:
    if t not in existing_templates:
        print(f"  [MISSING TEMPLATE BUG] Route renders '{t}', but file does not exist in templates/!")

# Collect all endpoints
endpoints = set()
tree = ast.parse(app_src)
for node in ast.walk(tree):
    if isinstance(node, ast.FunctionDef):
        for dec in node.decorator_list:
            if isinstance(dec, ast.Call) and getattr(dec.func, 'attr', '') == 'route':
                endpoints.add(node.name)
                for kw in dec.keywords:
                    if kw.arg == 'endpoint' and isinstance(kw.value, ast.Constant):
                        endpoints.add(kw.value.value)

for t_file in existing_templates:
    if not t_file.endswith('.html'):
        continue
    with open(os.path.join('templates', t_file), 'r', encoding='utf-8', errors='ignore') as f:
        t_code = f.read()
    # Check url_for
    for m in re.finditer(r'url_for\s*\(\s*[\'\"]([^\'\"]+)[\'\"]', t_code):
        target = m.group(1)
        if target != 'static' and target not in endpoints:
            print(f"  [BROKEN URL_FOR BUG] In template '{t_file}': url_for('{target}') does not exist in Flask endpoints!")

# Check Jinja syntax compilation
import jinja2
jinja_env = jinja2.Environment(loader=jinja2.FileSystemLoader('templates'))
for t_file in existing_templates:
    if not t_file.endswith('.html'):
        continue
    try:
        jinja_env.get_template(t_file)
    except Exception as e:
        print(f"  [JINJA SYNTAX ERROR] In template '{t_file}': {e}")

# -------------------------------------------------------------
# 4. EXAM ENGINE LOGIC & EDGE CASES AUDIT
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 4: Exam Engine Logic Tests")

# Test calculate_score
try:
    # Test scoring logic
    print("  Testing calculate_score...")
    # Setup mock data in mem_db
    c.execute("INSERT INTO students (student_id, name, class, subject, status) VALUES ('TEST001', 'Test Student', '10', 'Science', 'In Progress')")
    c.execute("INSERT INTO questions (id, class, subject, question, option_a, option_b, option_c, option_d, correct_answer) VALUES (1, '10', 'Science', 'Q1', 'A', 'B', 'C', 'D', 'A')")
    c.execute("INSERT INTO questions (id, class, subject, question, option_a, option_b, option_c, option_d, correct_answer) VALUES (2, '10', 'Science', 'Q2', 'A', 'B', 'C', 'D', 'B')")
    c.execute("INSERT INTO responses (student_id, question_id, selected_option) VALUES ('TEST001', 1, 'A')") # correct
    c.execute("INSERT INTO responses (student_id, question_id, selected_option) VALUES ('TEST001', 2, 'C')") # wrong
    mem_db.commit()

    score, total, pct = app.calculate_score('TEST001', 'Science', negative_marking=0, negative_value=0.33)
    print(f"  calculate_score without negative marking: score={score}, total={total}, pct={pct}")
    score_neg, total_neg, pct_neg = app.calculate_score('TEST001', 'Science', negative_marking=1, negative_value=0.33)
    print(f"  calculate_score with negative marking: score={score_neg}, total={total_neg}, pct={pct_neg}")
except Exception as e:
    print(f"  [SCORING BUG] calculate_score threw an exception: {e}")
    traceback.print_exc()

# -------------------------------------------------------------
# 5. SAFE JSON PARSER & LATEX TESTS
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 5: safe_json_loads & AI Output Recovery Tests")

test_cases = [
    ('Standard JSON', '{"question": "What is 2+2?", "option_a": "4", "correct_answer": "A"}'),
    ('Markdown fenced JSON', '```json\n[{"question": "Q1", "correct_answer": "A"}]\n```'),
    ('LaTeX with \\frac', '[{"question": "Find $\\frac{a}{b}$", "option_a": "\\frac{1}{2}", "correct_answer": "A"}]'),
    ('Raw Formfeed \\f in \\frac', '[\n  {\n    "question": "Solve $\x0crac{x}{2} = 5$",\n    "option_a": "\x0crac{1}{2}",\n    "correct_answer": "A"\n  }\n]'),
    ('Raw Tab \\t in \\tan and \\text', '[\n  {\n    "question": "Value of $\x09an(45^\\circ)$ is:",\n    "option_a": "\x09ext{one}",\n    "correct_answer": "A"\n  }\n]'),
    ('Truncated JSON cut off mid-stream', '[{"question": "Cut off question", "option_a": "Option A", "option_b": "Option B"'),
    ('Trailing commas before close', '[{"question": "Trailing comma", "option_a": "A", "option_b": "B",},]'),
    ('Single quotes Python style dict', "[{'question': 'Single quote dict', 'correct_answer': 'A'}]")
]

for name, payload in test_cases:
    try:
        res = app.safe_json_loads(payload)
        status = "PASSED" if res is not None else "FAILED (Returned None)"
        print(f"  safe_json_loads [{name}]: {status}")
    except Exception as e:
        print(f"  [SAFE_JSON_LOADS BUG] Failed on [{name}]: {e}")

# -------------------------------------------------------------
# 6. EXAM SHUFFLING & RESPONSE MAPPING AUDIT
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 6: Question Shuffling & Option Permutation Logic")

try:
    # Test shuffle_questions_for_student
    q_ids = [1, 2]
    app.shuffle_questions_for_student('TEST001', q_ids)
    shuffled = app.get_shuffled_questions('TEST001', q_ids)
    print(f"  shuffled_questions returned {len(shuffled)} questions.")
except Exception as e:
    print(f"  [SHUFFLING BUG] Shuffling logic failed: {e}")
    traceback.print_exc()

# -------------------------------------------------------------
# 7. EXPORT FUNCTIONS (DOCX, PDF, EXCEL) AUDIT
# -------------------------------------------------------------
print("\n>>> AUDIT SECTION 7: Document Export Handlers Audit")

try:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test"
    print("  openpyxl workbook creation OK.")
except Exception as e:
    print(f"  [EXCEL EXPORT BUG] openpyxl error: {e}")

try:
    import docx
    doc = docx.Document()
    doc.add_heading("Test Heading", level=1)
    print("  python-docx document creation OK.")
except Exception as e:
    print(f"  [DOCX EXPORT BUG] docx error: {e}")

try:
    import schemdraw
    with schemdraw.Drawing(show=False) as d:
        pass
    print("  schemdraw drawing creation OK.")
except Exception as e:
    print(f"  [CIRCUIT DRAWING BUG] schemdraw error: {e}")

print("\n" + "=" * 60)
print("AUDIT SCAN COMPLETE.")
print("=" * 60)
