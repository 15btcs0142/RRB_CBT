import os
import time
import csv
import requests
from dotenv import load_dotenv
import sqlite3
import datetime
import uuid
import random
import base64
import hashlib
import json
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file, make_response, has_request_context
from werkzeug.utils import secure_filename
import openpyxl
from weasyprint import HTML, CSS
# ─── WeasyPrint 61+ compatibility fix ──────────────────────────────
import weasyprint.pdf
import weasyprint.pdf.stream
import re
import ast
import image_handler
import exemplar_extractor
import curriculum_data
import question_bank
from validators import validate_schema, STUDENT_LOGIN_SCHEMA, ADMIN_LOGIN_SCHEMA, SAVE_ANSWER_SCHEMA, REATTEMPT_REQUEST_SCHEMA

def safe_json_loads(text):
    r"""
    Ultra-resilient JSON parser for AI-generated question papers.
    Handles:
      1. Markdown fences & wrapping text.
      2. LaTeX commands (\frac, \Omega, \mu, \times, \Delta, \theta, etc.).
      3. Unescaped quotes inside JSON string values.
      4. Literal unescaped newlines/tabs inside string values.
      5. Trailing commas before ] and }.
      6. Truncated / cut-off JSON (auto-closes braces and brackets).
      7. Python dict syntax (single quotes, True/False/None).
      8. Regex-based heuristic structural recovery fallback.
    """
    if not text:
        return None

    # Step 1: Strip outer markdown code fences
    s = re.sub(r'^\s*```(?:json|JSON)?\s*\n?', '', text)
    s = re.sub(r'\n?```\s*$', '', s)
    s = s.strip()

    # Step 2: Extract JSON candidate { ... } or [ ... ]
    match = re.search(r'(\{.*\}|\[.*\])', s, re.DOTALL)
    candidate = match.group(1) if match else s

    # Helper: repair single backslashes in LaTeX
    def repair_latex(val):
        if not val:
            return val
        # Fix control characters that got converted by python/json
        val = re.sub(r'\x0c(?=rac|orall|orce|lat|rown)', r'\\f', val)
        val = re.sub(r'\x09(?=ext|imes|heta|au|an|o|op|riangle|ilde)', r'\\t', val)
        val = re.sub(r'\x0d(?=ho|ight|ad|eal)', r'\\r', val)
        val = re.sub(r'\x08(?=eta|ar|inom|egin)', r'\\b', val)

        def _fix(m):
            cmd = m.group(1)
            if cmd in ('"', '\\', '/'):
                return '\\' + cmd
            if cmd in ('n', 'r', 't', 'b', 'f'):
                return '\\' + cmd
            if cmd.startswith('u') and len(cmd) == 5 and all(c in '0123456789abcdefABCDEF' for c in cmd[1:]):
                return '\\' + cmd
            return '\\\\' + cmd

        pattern = r'(?<!\\)\\((\(|\)|\[|\]|\{|\}|[a-zA-Z]+|[+\-=\-><\$%\&#\^~\'"]))'
        return re.sub(pattern, _fix, val)

    # Helper: auto-close truncated JSON
    def auto_close_json(raw_json):
        t = raw_json.strip()
        stack = []
        in_string = False
        escape = False
        cleaned_chars = []
        for char in t:
            if in_string:
                if escape:
                    escape = False
                elif char == '\\':
                    escape = True
                elif char == '"':
                    in_string = False
            else:
                if char == '"':
                    in_string = True
                elif char in ('{', '['):
                    stack.append(char)
                elif char == '}':
                    if stack and stack[-1] == '{': stack.pop()
                elif char == ']':
                    if stack and stack[-1] == '[': stack.pop()
            cleaned_chars.append(char)

        repaired = "".join(cleaned_chars)
        if in_string:
            repaired += '"'
        repaired = re.sub(r',\s*$', '', repaired.strip())
        while stack:
            opener = stack.pop()
            if opener == '{':
                repaired = re.sub(r',\s*$', '', repaired.strip()) + '}'
            elif opener == '[':
                repaired = re.sub(r',\s*$', '', repaired.strip()) + ']'
        return repaired

    # Helper: fix unescaped internal quotes inside JSON string fields
    def fix_unescaped_internal_quotes(json_str):
        def _fix_line(m):
            prefix = m.group(1)
            content = m.group(2)
            suffix = m.group(3)
            escaped_content = re.sub(r'(?<!\\)"', r'\"', content)
            return f'{prefix}{escaped_content}{suffix}'

        line_pattern = r'("(?:question|instruction|section_title|section_label|exam_type|subject|title|problem_statement|input_format|output_format|constraints|sample_input|sample_output|input|expected_output)":\s*")(.*?)("(?:\s*,\s*|\s*\}|\s*\]|\s*$))'
        return re.sub(line_pattern, _fix_line, json_str, flags=re.DOTALL)

    # Strategy 1: Direct JSON parse with strict=False
    try:
        return json.loads(candidate, strict=False)
    except Exception:
        pass

    # Strategy 2: LaTeX repaired
    repaired_s = candidate
    try:
        repaired_s = repair_latex(candidate)
        return json.loads(repaired_s, strict=False)
    except Exception:
        pass

    # Strategy 3: Clean trailing commas & LaTeX repair
    clean_s = repaired_s
    try:
        clean_s = re.sub(r',\s*([\]}])', r'\1', repaired_s)
        return json.loads(clean_s, strict=False)
    except Exception:
        pass

    # Strategy 4: Fix unescaped internal quotes + LaTeX repair
    try:
        fixed_quotes = fix_unescaped_internal_quotes(clean_s)
        return json.loads(fixed_quotes, strict=False)
    except Exception:
        pass

    # Strategy 5: Auto-close truncated JSON
    try:
        closed = auto_close_json(clean_s)
        return json.loads(closed, strict=False)
    except Exception:
        pass

    try:
        closed_q = auto_close_json(fix_unescaped_internal_quotes(clean_s))
        return json.loads(closed_q, strict=False)
    except Exception:
        pass

    # Strategy 6: ast.literal_eval for Python-like representations
    try:
        py_literal = clean_s.replace('null', 'None').replace('true', 'True').replace('false', 'False')
        res = ast.literal_eval(py_literal)
        if isinstance(res, (dict, list)):
            return res
    except Exception:
        pass

    # Strategy 7: Fallback Structural Regex Extraction (recovers questions if JSON is severely mangled)
    try:
        sections_found = []
        sec_blocks = re.findall(r'\{\s*"section_label"\s*:\s*"([^"]+)"\s*,\s*"section_title"\s*:\s*"([^"]+)"(.*?)(?=\{\s*"section_label"|\Z)', s, re.DOTALL)
        for label, title, body in sec_blocks:
            questions_found = []
            q_matches = re.findall(r'\{\s*"number"\s*:\s*(\d+)\s*,\s*"question"\s*:\s*"(.*?)"\s*,\s*"options"\s*:\s*\[(.*?)\]', body, re.DOTALL)
            for q_num, q_text, opts_raw in q_matches:
                opts = [o.strip(' "\'') for o in opts_raw.split(',') if o.strip(' "\'')]
                questions_found.append({
                    "number": int(q_num),
                    "question": q_text.strip(),
                    "options": opts,
                    "sub_questions": [],
                    "needs_image": False,
                    "image_category": "none",
                    "circuit_type": None,
                    "params": {},
                    "image_keywords": []
                })
            if questions_found:
                sections_found.append({
                    "section_label": label,
                    "section_title": title,
                    "marks_per_question": 1,
                    "instruction": "Answer all questions.",
                    "questions": questions_found
                })
        if sections_found:
            return {"sections": sections_found}
    except Exception:
        pass

    return None

def parse_question_csv_row(row):
    """
    Flexible CSV parser supporting column headers:
    Questions, option a, option b, option c, option d, correct option
    (and all common variations like option_a, question, correct_answer).
    """
    if not isinstance(row, dict):
        return None

    # Case-insensitive normalized map
    norm = {}
    for k, v in row.items():
        if k is not None and str(k).strip():
            k_clean = str(k).strip().lower().replace('_', ' ')
            norm[k_clean] = str(v).strip() if v is not None else ''

    def get_val(*keys):
        for k in keys:
            k_clean = k.strip().lower().replace('_', ' ')
            if k_clean in norm and norm[k_clean]:
                return norm[k_clean]
        return ''

    q = get_val('questions', 'question', 'question text', 'q')
    if not q:
        return None

    opt_a = get_val('option a', 'optiona', 'a', 'option 1', 'option1')
    opt_b = get_val('option b', 'optionb', 'b', 'option 2', 'option2')
    opt_c = get_val('option c', 'optionc', 'c', 'option 3', 'option3')
    opt_d = get_val('option d', 'optiond', 'd', 'option 4', 'option4')

    raw_ans = get_val('correct option', 'correct answer', 'correct_answer', 'correctoption', 'answer', 'correct')
    
    # Normalize correct_answer to option_a, option_b, option_c, option_d
    ans_clean = raw_ans.lower().replace('_', ' ').strip()
    correct_ans = 'option_a'
    if ans_clean in ('option a', 'optiona', 'a', '1', 'option 1', 'option1'):
        correct_ans = 'option_a'
    elif ans_clean in ('option b', 'optionb', 'b', '2', 'option 2', 'option2'):
        correct_ans = 'option_b'
    elif ans_clean in ('option c', 'optionc', 'c', '3', 'option 3', 'option3'):
        correct_ans = 'option_c'
    elif ans_clean in ('option d', 'optiond', 'd', '4', 'option 4', 'option4'):
        correct_ans = 'option_d'
    elif raw_ans and opt_a and raw_ans == opt_a:
        correct_ans = 'option_a'
    elif raw_ans and opt_b and raw_ans == opt_b:
        correct_ans = 'option_b'
    elif raw_ans and opt_c and raw_ans == opt_c:
        correct_ans = 'option_c'
    elif raw_ans and opt_d and raw_ans == opt_d:
        correct_ans = 'option_d'
    elif raw_ans:
        correct_ans = raw_ans

    q_type = get_val('question type', 'type') or ('MCQ' if (opt_a or opt_b) else 'Short')
    marks = get_val('marks', 'mark') or '1'
    neg_mark = get_val('negative mark', 'negative marks') or '0.33'

    return {
        'question': q,
        'option_a': opt_a,
        'option_b': opt_b,
        'option_c': opt_c,
        'option_d': opt_d,
        'correct_answer': correct_ans,
        'question_type': q_type,
        'marks': marks,
        'negative_mark': neg_mark
    }

def _normalize_correct_answer(correct_raw, opt_a='', opt_b='', opt_c='', opt_d=''):
    """
    Normalize any correct_answer representation ('A', 'a', 'option_a', 'Option A', '1', or actual option text)
    to a standard uppercase letter ('A', 'B', 'C', 'D').
    """
    if not correct_raw:
        return ''
    c_str = str(correct_raw).strip()
    c_upper = c_str.upper()

    # 1. Single letter 'A', 'B', 'C', 'D'
    if c_upper in ('A', 'B', 'C', 'D'):
        return c_upper

    # 2. Number string '1', '2', '3', '4'
    if c_str in ('1', '2', '3', '4'):
        return chr(ord('A') + int(c_str) - 1)

    # 3. 'OPTION_A', 'OPTION A', 'OPT_A', 'OPTION_B', etc.
    c_clean = c_upper.replace(' ', '_')
    if c_clean in ('OPTION_A', 'OPT_A'):
        return 'A'
    if c_clean in ('OPTION_B', 'OPT_B'):
        return 'B'
    if c_clean in ('OPTION_C', 'OPT_C'):
        return 'C'
    if c_clean in ('OPTION_D', 'OPT_D'):
        return 'D'

    # 4. Text match against option_a, option_b, option_c, option_d
    c_lower = c_str.lower()
    if opt_a and str(opt_a).strip().lower() == c_lower:
        return 'A'
    if opt_b and str(opt_b).strip().lower() == c_lower:
        return 'B'
    if opt_c and str(opt_c).strip().lower() == c_lower:
        return 'C'
    if opt_d and str(opt_d).strip().lower() == c_lower:
        return 'D'

    return ''

def _save_or_update_result(c, student_id, name, class_, section, subject, score, total_questions, percentage, chapter, section_breakdown=None):
    """Save new result or update existing result row for (student_id, subject, chapter), optionally storing section_breakdown JSON."""
    c.execute("""SELECT id FROM results 
                 WHERE student_id=? AND LOWER(subject)=LOWER(?) AND (chapter=? OR (chapter IS NULL AND ?=''))""",
              (student_id, subject, chapter or '', chapter or ''))
    existing = c.fetchone()
    if existing:
        if section_breakdown is not None:
            c.execute("""UPDATE results 
                         SET name=?, class=?, section=?, score=?, total_questions=?, percentage=?, section_breakdown=?, test_date=CURRENT_TIMESTAMP
                         WHERE id=?""",
                      (name, class_, section or '', round(score, 2), total_questions, percentage, section_breakdown, existing['id']))
        else:
            c.execute("""UPDATE results 
                         SET name=?, class=?, section=?, score=?, total_questions=?, percentage=?, test_date=CURRENT_TIMESTAMP
                         WHERE id=?""",
                      (name, class_, section or '', round(score, 2), total_questions, percentage, existing['id']))
    else:
        c.execute("""INSERT INTO results 
                     (student_id, name, class, section, subject, score, total_questions, percentage, chapter, section_breakdown, test_date)
                     VALUES (?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)""",
                  (student_id, name, class_, section or '', subject, round(score, 2), total_questions, percentage, chapter or '', section_breakdown))

def _get_multi_section_info(class_, subject, test_no, conn):
    """Return test paper & config dict if (class_, subject, test_no) corresponds to an active Multi-Section Test."""
    if not (class_ and subject and test_no):
        return None
    c = conn.cursor()
    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    c.execute("""SELECT tp.id as tp_id, tp.subject as title, tp.test_no, tp.class,
                        tsc.timing_mode, tsc.shuffle_sections, tsc.two_stage_instructions
                 FROM test_papers tp
                 JOIN test_section_config tsc ON tp.id = tsc.test_paper_id
                 WHERE (tp.class=? OR tp.class LIKE ?) AND LOWER(tp.subject)=LOWER(?) AND tp.test_no=? AND tp.paper_type='multi_section'
                 LIMIT 1""",
              (class_, f"%{clean_cls}%", subject, test_no))
    row = c.fetchone()
    return dict(row) if row else None

def get_working_model(api_key):
    """
    Try to find a working Gemini model by listing available models.
    Returns a model name string or default 'gemini-3.5-flash-lite'.
    """
    import urllib.request, json
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            models = data.get('models', [])
            supported_models = []
            for model in models:
                name = model.get('name', '').replace('models/', '')
                supported = model.get('supportedGenerationMethods', [])
                if 'generateContent' in supported:
                    supported_models.append(name)

            # Preference priority for valid active models
            for preferred in ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-flash-latest', 'gemini-2.5-pro', 'gemini-flash-lite-latest']:
                if preferred in supported_models:
                    return preferred

            for name in supported_models:
                if 'flash' in name or 'pro' in name:
                    return name

            if supported_models:
                return supported_models[0]
    except Exception:
        pass
    return 'gemini-3.5-flash-lite'

def call_gemini_generate_content(prompt, api_key, generation_config=None, timeout=60):
    """
    Robust Gemini API caller with automatic candidate model fallbacks and HTTP 429 backoff/retry.
    Returns tuple: (result_dict, error_message)
    """
    import urllib.request, urllib.error, json, time
    if not generation_config:
        generation_config = {"temperature": 0.7, "maxOutputTokens": 8192}
    else:
        generation_config = dict(generation_config)

    # Force JSON output mode if prompt requests JSON structure
    if ("JSON" in prompt or "json" in prompt) and "response_mime_type" not in generation_config:
        generation_config["response_mime_type"] = "application/json"

    candidate_models = ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-flash-latest', 'gemini-2.5-pro', 'gemini-flash-lite-latest']
    detected = get_working_model(api_key)
    if detected and detected not in candidate_models:
        candidate_models.insert(0, detected)

    payload_bytes = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": generation_config
    }).encode('utf-8')

    last_error = ""

    for model_name in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        for attempt in range(2):
            try:
                req = urllib.request.Request(
                    url,
                    data=payload_bytes,
                    headers={'Content-Type': 'application/json'},
                    method='POST'
                )
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    resp_data = resp.read().decode('utf-8')
                    if resp_data.lstrip().startswith('<'):
                        last_error = "Gemini API returned HTML instead of JSON."
                        break
                    return json.loads(resp_data), None
            except urllib.error.HTTPError as e:
                err_body = e.read().decode('utf-8')
                last_error = f"{model_name} HTTP {e.code}: {err_body}"
                if e.code in (404, 400):
                    break
                elif e.code == 429:
                    time.sleep(4)
                    continue
                else:
                    break
            except Exception as e:
                last_error = str(e)
                break

    if "429" in last_error or "quota" in last_error.lower() or "resource_exhausted" in last_error.lower():
        clean_msg = "Gemini API rate limit / free tier quota exceeded (HTTP 429). Please wait 30–60 seconds and try again, or check your API key quota at https://ai.google.dev."
    else:
        clean_msg = f"Gemini API error: {last_error[:300]}"

    return None, clean_msg

def _get_ai_keys():
    """
    Parse apikey.env, .api_key, and OS environment for API keys across providers:
    Gemini, DeepSeek, OpenAI (ChatGPT), and Claude (Anthropic).
    Returns dict: {'gemini': '...', 'deepseek': '...', 'openai': '...', 'claude': '...'}
    """
    keys = {
        'gemini': os.environ.get('GEMINI_API_KEY', '').strip().strip('"\''),
        'deepseek': os.environ.get('DEEPSEEK_API_KEY', '').strip().strip('"\''),
        'openai': os.environ.get('OPENAI_API_KEY', '').strip().strip('"\''),
        'claude': os.environ.get('CLAUDE_API_KEY', '') or os.environ.get('ANTHROPIC_API_KEY', ''),
        'huggingface': os.environ.get('HUGGINGFACE_API_KEY', '') or os.environ.get('HF_TOKEN', '') or os.environ.get('HF_API_KEY', ''),
        'hf_model': os.environ.get('HF_MODEL', '') or os.environ.get('HUGGINGFACE_MODEL', '')
    }
    for k in keys:
        keys[k] = keys[k].strip().strip('"\'')

    # 1. Primary Master Source: apikey.env file
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'apikey.env')
    if os.path.exists(env_path):
        try:
            with open(env_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()

            is_single = True
            for line in lines:
                line_s = line.strip()
                if not line_s or line_s.startswith('#'):
                    continue
                if '=' in line_s:
                    is_single = False
                    parts = line_s.split('=', 1)
                    k_name = parts[0].strip().upper()
                    val = parts[1].strip().strip('"\'')
                    if 'GEMINI' in k_name and val:
                        keys['gemini'] = val
                    elif 'DEEPSEEK' in k_name and val:
                        keys['deepseek'] = val
                    elif ('OPENAI' in k_name or 'CHATGPT' in k_name) and val:
                        keys['openai'] = val
                    elif ('CLAUDE' in k_name or 'ANTHROPIC' in k_name) and val:
                        keys['claude'] = val
                    elif ('HUGGINGFACE_MODEL' in k_name or 'HF_MODEL' in k_name) and val:
                        keys['hf_model'] = val
                    elif ('HUGGINGFACE' in k_name or 'HF_' in k_name) and val:
                        keys['huggingface'] = val

            if is_single and lines:
                raw_val = ''.join(lines).strip().strip('"\'')
                if raw_val and not raw_val.startswith('#'):
                    keys['gemini'] = raw_val
        except Exception:
            pass

    # 2. Secondary Fallback / Sync: .api_key file (used by batch tools)
    dot_key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.api_key')
    if os.path.exists(dot_key_path):
        try:
            val = open(dot_key_path, 'r', encoding='utf-8').read().strip().strip('"\'')
            if val and not val.startswith('#') and not keys['gemini']:
                keys['gemini'] = val
        except Exception:
            pass

    # Keep .api_key in sync with current active Gemini key
    if keys['gemini']:
        try:
            with open(dot_key_path, 'w', encoding='utf-8') as f:
                f.write(keys['gemini'])
        except Exception:
            pass

    return keys

def _call_deepseek_api(prompt, api_key, timeout=90):
    """Call DeepSeek AI chat completions API (OpenAI-compatible)"""
    import urllib.request, urllib.error, json
    url = "https://api.deepseek.com/chat/completions"
    payload = json.dumps({
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 4096
    }).encode('utf-8')
    req = urllib.request.Request(
        url, data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method='POST'
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            res = json.loads(resp.read().decode('utf-8'))
            text = res['choices'][0]['message']['content'].strip()
            return text, None
    except urllib.error.HTTPError as e:
        err_body = e.read().decode('utf-8')
        try:
            msg = json.loads(err_body).get('error', {}).get('message', err_body)
        except Exception:
            msg = err_body
        return None, f"DeepSeek HTTP {e.code}: {msg[:300]}"
    except Exception as e:
        return None, f"DeepSeek Error: {str(e)}"

def _call_openai_api(prompt, api_key, timeout=90):
    """Call OpenAI ChatGPT API (gpt-4o-mini / gpt-3.5-turbo)"""
    import urllib.request, urllib.error, json
    models = ["gpt-4o-mini", "gpt-3.5-turbo"]
    last_err = ""
    for model_name in models:
        url = "https://api.openai.com/v1/chat/completions"
        payload = json.dumps({
            "model": model_name,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 4096
        }).encode('utf-8')
        req = urllib.request.Request(
            url, data=payload,
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
            method='POST'
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                res = json.loads(resp.read().decode('utf-8'))
                text = res['choices'][0]['message']['content'].strip()
                return text, None
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8')
            try:
                msg = json.loads(err_body).get('error', {}).get('message', err_body)
            except Exception:
                msg = err_body
            last_err = f"OpenAI HTTP {e.code}: {msg[:300]}"
            if e.code in (404, 400) and 'model' in last_err.lower():
                continue
            else:
                break
        except Exception as e:
            last_err = f"OpenAI Error: {str(e)}"
            break
    return None, last_err

def _call_claude_api(prompt, api_key, timeout=90):
    """Call Anthropic Claude API (claude-3-5-haiku-20241022 / claude-3-haiku-20240307)"""
    import urllib.request, urllib.error, json
    models = ["claude-3-5-haiku-20241022", "claude-3-haiku-20240307"]
    last_err = ""
    for model_name in models:
        url = "https://api.anthropic.com/v1/messages"
        payload = json.dumps({
            "model": model_name,
            "max_tokens": 4096,
            "messages": [{"role": "user", "content": prompt}]
        }).encode('utf-8')
        req = urllib.request.Request(
            url, data=payload,
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01"
            },
            method='POST'
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                res = json.loads(resp.read().decode('utf-8'))
                text = res['content'][0]['text'].strip()
                return text, None
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8')
            try:
                err_json = json.loads(err_body)
                msg = err_json.get('error', {}).get('message', err_body)
            except Exception:
                msg = err_body
            last_err = f"Claude HTTP {e.code}: {msg[:300]}"
            if e.code in (404, 400) and 'model' in last_err.lower():
                continue
            else:
                break
        except Exception as e:
            last_err = f"Claude Error: {str(e)}"
            break
    return None, last_err

def _call_huggingface_api(prompt, api_key, timeout=90):
    """Call Hugging Face Inference API (Serverless Open-Source models)"""
    import urllib.request, urllib.error, json
    keys = _get_ai_keys()
    custom_model = keys.get('hf_model')
    models = [
        "Qwen/Qwen2.5-7B-Instruct",
        "meta-llama/Llama-3.1-8B-Instruct",
        "mistralai/Mistral-7B-Instruct-v0.3",
        "Qwen/Qwen2.5-Coder-32B-Instruct"
    ]
    if custom_model and custom_model not in models:
        models.insert(0, custom_model)
    elif custom_model and custom_model in models:
        models.remove(custom_model)
        models.insert(0, custom_model)
    last_err = ""
    for model_name in models:
        url = "https://router.huggingface.co/v1/chat/completions"
        payload = json.dumps({
            "model": model_name,
            "messages": [
                {"role": "system", "content": "You are a professional educational question paper generator. Always return clean JSON array output."},
                {"role": "user", "content": prompt}
            ],
            "max_tokens": 4096,
            "temperature": 0.7
        }).encode('utf-8')
        req = urllib.request.Request(
            url, data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) RRB-CBT-Engine"
            },
            method='POST'
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                res = json.loads(resp.read().decode('utf-8'))
                text = res['choices'][0]['message']['content'].strip()
                return text, None
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8')
            try:
                err_json = json.loads(err_body)
                msg = err_json.get('error', {}).get('message', err_body)
                if isinstance(err_json.get('error'), str):
                    msg = err_json.get('error')
            except Exception:
                msg = err_body
            last_err = f"Hugging Face HTTP {e.code}: {msg[:300]}"
            if e.code in (404, 400, 503):
                continue
            else:
                break
        except Exception as e:
            last_err = f"Hugging Face Error: {str(e)}"
            break
    return None, last_err

def generate_ai_content(prompt, timeout=90):
    """
    Multi-provider AI generator with automatic failover across Gemini, DeepSeek, OpenAI, Claude, and Hugging Face.
    Returns (result_dict, provider_name, error_message).
    """
    keys = _get_ai_keys()
    errors = []

    # Priority 1: Google Gemini
    if keys.get('gemini'):
        res_dict, err = call_gemini_generate_content(prompt, keys['gemini'], timeout=timeout)
        if res_dict:
            return res_dict, 'Gemini AI', None
        app.logger.warning(f"Gemini AI failed, trying alternate providers: {err}")
        errors.append(f"Gemini: {err}")

    # Priority 2: DeepSeek AI
    if keys.get('deepseek'):
        text, err = _call_deepseek_api(prompt, keys['deepseek'], timeout=timeout)
        if text:
            res_dict = {'candidates': [{'content': {'parts': [{'text': text}]}}]}
            return res_dict, 'DeepSeek AI', None
        app.logger.warning(f"DeepSeek AI failed: {err}")
        errors.append(f"DeepSeek: {err}")

    # Priority 3: OpenAI ChatGPT
    if keys.get('openai'):
        text, err = _call_openai_api(prompt, keys['openai'], timeout=timeout)
        if text:
            res_dict = {'candidates': [{'content': {'parts': [{'text': text}]}}]}
            return res_dict, 'ChatGPT (OpenAI)', None
        app.logger.warning(f"OpenAI ChatGPT failed: {err}")
        errors.append(f"OpenAI: {err}")

    # Priority 4: Anthropic Claude
    if keys.get('claude'):
        text, err = _call_claude_api(prompt, keys['claude'], timeout=timeout)
        if text:
            res_dict = {'candidates': [{'content': {'parts': [{'text': text}]}}]}
            return res_dict, 'Claude AI', None
        app.logger.warning(f"Claude AI failed: {err}")
        errors.append(f"Claude: {err}")

    # Priority 5: Hugging Face AI
    if keys.get('huggingface'):
        text, err = _call_huggingface_api(prompt, keys['huggingface'], timeout=timeout)
        if text:
            res_dict = {'candidates': [{'content': {'parts': [{'text': text}]}}]}
            return res_dict, 'Hugging Face AI', None
        app.logger.warning(f"Hugging Face AI failed: {err}")
        errors.append(f"Hugging Face: {err}")

    if not any(keys.values()):
        return None, None, "No AI API key found. Please set GEMINI_API_KEY, DEEPSEEK_API_KEY, OPENAI_API_KEY, or CLAUDE_API_KEY in apikey.env or .api_key."

    combined_err = " | ".join(errors) if errors else "All configured AI providers failed."
    return None, None, f"AI Generation Failed: {combined_err}"
# ─── WeasyPrint & pydyf 0.12+ compatibility fixes ──────────────────────────
import pydyf

# 1. Patch pydyf.PDF.__init__ compatibility with WeasyPrint 60.0 & pydyf 0.12+
_orig_pydyf_pdf_init = pydyf.PDF.__init__

def _patched_pydyf_pdf_init(self, version='1.7', identifier=None, *args, **kwargs):
    _orig_pydyf_pdf_init(self)
    if version:
        self.version = version.encode('ascii') if isinstance(version, str) else version
    else:
        self.version = b'1.7'
    if identifier:
        self.identifier = identifier.encode('ascii') if isinstance(identifier, str) else identifier

pydyf.PDF.__init__ = _patched_pydyf_pdf_init

# 2. Patch Stream.transform compatibility for WeasyPrint 60+ stream
def _patched_stream_transform(self, a=1, b=0, c=0, d=1, e=0, f=0):
    if hasattr(self, 'set_matrix'):
        self.set_matrix(a, b, c, d, e, f)
    else:
        self._matrix = (a, b, c, d, e, f)

weasyprint.pdf.stream.Stream.transform = _patched_stream_transform

# 3. Patch Stream.text_matrix alias compatibility with pydyf 0.12+
if not hasattr(weasyprint.pdf.stream.Stream, 'text_matrix'):
    def _stream_text_matrix(self, a=1, b=0, c=0, d=1, e=0, f=0):
        if hasattr(self, 'set_matrix'):
            self.set_matrix(a, b, c, d, e, f)
        elif hasattr(self, 'set_text_matrix'):
            self.set_text_matrix(a, b, c, d, e, f)
        else:
            self._matrix = (a, b, c, d, e, f)
    
    weasyprint.pdf.stream.Stream.text_matrix = _stream_text_matrix

load_dotenv('.env')
load_dotenv('apikey.env')

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY') or os.environ.get('FLASK_SECRET_KEY') or 'rrb-cbt-v104-secret-key-2024'
app.config['TEMPLATES_AUTO_RELOAD'] = True

# ─── JWT Authentication Subsystem ──────────────────────────────────────
import jwt
from flask import g

JWT_SECRET_KEY = os.environ.get('JWT_SECRET_KEY', 'rrb_cbt_jwt_secret_key_2026_x98f72a')
JWT_EXPIRATION_HOURS = int(os.environ.get('JWT_EXPIRATION_HOURS', 24))

def generate_jwt_token(identity, role, name=None, expires_in_hours=None):
    """
    Generate a signed JWT token using PyJWT.
    Claims: sub (identity), role (admin/teacher/student), name, iat, exp.
    """
    if expires_in_hours is None:
        expires_in_hours = JWT_EXPIRATION_HOURS
        
    now = datetime.datetime.now(datetime.timezone.utc)
    expiration = now + datetime.timedelta(hours=expires_in_hours)
    
    payload = {
        'sub': str(identity),
        'role': role,
        'name': name or str(identity),
        'iat': int(now.timestamp()),
        'exp': int(expiration.timestamp())
    }
    
    token = jwt.encode(payload, JWT_SECRET_KEY, algorithm='HS256')
    return token, int(expiration.timestamp())

def decode_jwt_token(token):
    """
    Decode and validate a signed JWT token using PyJWT.
    Returns payload dict if valid, or raises PyJWT exception if invalid/expired.
    """
    return jwt.decode(token, JWT_SECRET_KEY, algorithms=['HS256'])

def jwt_required(allowed_roles=None):
    """
    Decorator for API-only routes.
    Validates Authorization: Bearer <token> header.
    Attaches decoded payload to Flask context g.jwt_payload & g.jwt_user.
    Returns HTTP 401 Unauthorized if missing, expired, or invalid.
    """
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            auth_header = request.headers.get('Authorization')
            if not auth_header:
                return jsonify({'status': 'error', 'message': 'Missing Authorization header'}), 401
                
            parts = auth_header.split()
            if len(parts) != 2 or parts[0].lower() != 'bearer':
                return jsonify({'status': 'error', 'message': 'Authorization header format must be Bearer <token>'}), 401
                
            token = parts[1]
            try:
                payload = decode_jwt_token(token)
                g.jwt_payload = payload
                g.jwt_user = {
                    'id': payload.get('sub'),
                    'role': payload.get('role'),
                    'name': payload.get('name')
                }
                
                if allowed_roles:
                    roles_list = [allowed_roles] if isinstance(allowed_roles, str) else allowed_roles
                    if payload.get('role') not in roles_list:
                        return jsonify({'status': 'error', 'message': 'Forbidden: Insufficient privileges'}), 403
                        
            except jwt.ExpiredSignatureError:
                return jsonify({'status': 'error', 'message': 'Token has expired'}), 401
            except jwt.InvalidTokenError as e:
                return jsonify({'status': 'error', 'message': f'Invalid token: {str(e)}'}), 401
                
            return f(*args, **kwargs)
        return decorated
    return decorator

# ─── Flask-Limiter Setup ───────────────────────────────────────────
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[],
    storage_uri="memory://"
)

@app.errorhandler(429)
def ratelimit_handler(e):
    retry_after = getattr(e, 'retry_after', 60) or 60
    return jsonify({
        'status': 'error',
        'rate_limit': True,
        'retry_after': int(retry_after),
        'message': f'Server rate limit reached (13 requests per minute). Please wait {int(retry_after)} seconds before trying again.'
    }), 429
# Configuration
UPLOAD_FOLDER = 'uploads'
EXPORT_FOLDER = 'exports'
QB_UPLOAD_FOLDER = 'question_bank_uploads'  # Raw PDFs for Question Bank
BRANDING_FOLDER = os.path.join('static', 'uploads', 'branding')
STUDENT_PIC_FOLDER = os.path.join('static', 'student_profile_pic')
TEACHER_PIC_FOLDER = os.path.join('static', 'Teacher_profile_picture')
ALLOWED_EXTENSIONS = {'csv'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)
os.makedirs(QB_UPLOAD_FOLDER, exist_ok=True)
os.makedirs('question_bank_images', exist_ok=True)
os.makedirs(os.path.join('static', 'uploads'), exist_ok=True)
os.makedirs(BRANDING_FOLDER, exist_ok=True)
os.makedirs(STUDENT_PIC_FOLDER, exist_ok=True)
os.makedirs(TEACHER_PIC_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXPORT_FOLDER'] = EXPORT_FOLDER
app.config['BRANDING_FOLDER'] = BRANDING_FOLDER
app.config['STUDENT_PIC_FOLDER'] = STUDENT_PIC_FOLDER
app.config['TEACHER_PIC_FOLDER'] = TEACHER_PIC_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload

# ── helpers ──────────────────────────────────────────────
def hash_password(password):
    """
    Hash a password using SHA256 encryption.
    """
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, hashed):
    """
    Verify if a provided password matches its hashed version.
    """
    return hash_password(password) == hashed

def save_student_picture(file, student_id):
    """
    Save a student profile picture to the file system and return the file path.
    """
    if not file or file.filename == '':
        return None
    if not allowed_image_file(file.filename):
        return None
    ext = file.filename.rsplit('.', 1)[1].lower()
    filename = f"student_{student_id}.{ext}"
    filepath = os.path.join(STUDENT_PIC_FOLDER, filename)
    # Remove old picture if different extension
    for old_ext in ALLOWED_IMAGE_EXTENSIONS:
        old_path = os.path.join(STUDENT_PIC_FOLDER, f"student_{student_id}.{old_ext}")
        if os.path.exists(old_path) and old_path != filepath:
            os.remove(old_path)
    file.save(filepath)
    return os.path.join('student_profile_pic', filename).replace('\\', '/')

def save_teacher_picture(file, teacher_id):
    """
    Save a teacher profile picture to the file system and return the file path.
    """
    if not file or file.filename == '':
        return None
    if not allowed_image_file(file.filename):
        return None
    ext = file.filename.rsplit('.', 1)[1].lower()
    filename = f"teacher_{teacher_id}.{ext}"
    filepath = os.path.join(TEACHER_PIC_FOLDER, filename)
    for old_ext in ALLOWED_IMAGE_EXTENSIONS:
        old_path = os.path.join(TEACHER_PIC_FOLDER, f"teacher_{teacher_id}.{old_ext}")
        if os.path.exists(old_path) and old_path != filepath:
            os.remove(old_path)
    file.save(filepath)
    return os.path.join('Teacher_profile_picture', filename).replace('\\', '/')

def allowed_image_file(filename):
    """
    Check if a filename has an allowed image extension (png, jpg, jpeg, gif, webp).
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS

def save_question_image(file, class_, subject, question_id):
    """
    Save a question image to the appropriate folder based on class and subject.
    """
    if not file or file.filename == '':
        return None
    if not allowed_image_file(file.filename):
        return None

    folder_name = f"{class_.strip()}_{subject.strip()}".replace(' ', '_')
    folder_path = os.path.join('static', 'uploads', folder_name)
    os.makedirs(folder_path, exist_ok=True)

    ext = file.filename.rsplit('.', 1)[1].lower()
    filename = f"{folder_name}_{question_id}.{ext}"
    filepath = os.path.join(folder_path, filename)

    if os.path.exists(filepath):
        os.remove(filepath)

    file.save(filepath)
    return os.path.join('uploads', folder_name, filename).replace('\\', '/')

# Database initialization
def init_db():
    """
    Initialize the SQLite database with all required tables and schema. Handles migrations from old schemas.
    """
    conn = sqlite3.connect('database.db', timeout=30.0)
    c = conn.cursor()
    c.execute("PRAGMA busy_timeout = 30000")
    c.execute("PRAGMA journal_mode=DELETE")

    c.execute('''CREATE TABLE IF NOT EXISTS students
                 (student_id TEXT PRIMARY KEY, name TEXT, class TEXT, subject TEXT, ip TEXT,
                  status TEXT DEFAULT 'Not Started', exam_started_at TIMESTAMP,
                  admission_no TEXT, section TEXT, dob TEXT, house TEXT, 
                  parents_name TEXT, address TEXT, picture TEXT)''')
    c.execute("PRAGMA table_info(students)")
    columns = [col[1] for col in c.fetchall()]
    if 'exam_started_at' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN exam_started_at TIMESTAMP")
    if 'admission_no' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN admission_no TEXT")
    if 'section' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN section TEXT")
    if 'dob' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN dob TEXT")
    if 'house' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN house TEXT")
    if 'parents_name' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN parents_name TEXT")
    if 'address' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN address TEXT")
    if 'picture' not in columns:
        c.execute("ALTER TABLE students ADD COLUMN picture TEXT")
    # Add test_no and system_name columns to students if not exists
    c.execute("PRAGMA table_info(students)")
    cols = [col[1] for col in c.fetchall()]
    if 'test_no' not in cols:
        c.execute("ALTER TABLE students ADD COLUMN test_no TEXT DEFAULT ''")
    if 'system_name' not in cols:
        c.execute("ALTER TABLE students ADD COLUMN system_name TEXT DEFAULT ''")

    c.execute('''CREATE TABLE IF NOT EXISTS questions
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, class TEXT, subject TEXT, question TEXT,
                  option_a TEXT, option_b TEXT, option_c TEXT, option_d TEXT, correct_answer TEXT,
                  image_path TEXT, chapter TEXT)''')
        # Add chapter column if not exists
    c.execute("PRAGMA table_info(questions)")
    columns = [col[1] for col in c.fetchall()]
    if 'chapter' not in columns:
        c.execute("ALTER TABLE questions ADD COLUMN chapter TEXT")
    if 'class' not in columns:
        c.execute("ALTER TABLE questions ADD COLUMN class TEXT")
    if 'subject' not in columns:
        c.execute("ALTER TABLE questions ADD COLUMN subject TEXT")
    if 'image_path' not in columns:
        c.execute("ALTER TABLE questions ADD COLUMN image_path TEXT")
    if 'section' not in columns:
        c.execute("ALTER TABLE questions ADD COLUMN section TEXT")
    if 'test_no' not in columns:
        c.execute("ALTER TABLE questions ADD COLUMN test_no TEXT")
    for hi_col in ['question_hi', 'option_a_hi', 'option_b_hi', 'option_c_hi', 'option_d_hi']:
        if hi_col not in columns:
            c.execute(f"ALTER TABLE questions ADD COLUMN {hi_col} TEXT")

    c.execute('''CREATE TABLE IF NOT EXISTS responses
                 (student_id TEXT, question_id INTEGER, selected_option TEXT,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  PRIMARY KEY (student_id, question_id))''')
    c.execute("PRAGMA table_info(responses)")
    resp_columns = [col[1] for col in c.fetchall()]
    if 'created_at' not in resp_columns:
        c.execute("ALTER TABLE responses ADD COLUMN created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP")

    c.execute('''CREATE TABLE IF NOT EXISTS results
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT, name TEXT, 
                  class TEXT, subject TEXT, score INTEGER, total_questions INTEGER,
                  percentage REAL, test_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  chapter TEXT)''')
    # Migrate old results table if it has the old schema (student_id was PRIMARY KEY)
    c.execute("PRAGMA table_info(results)")
    res_columns = [col[1] for col in c.fetchall()]
    if 'id' not in res_columns:
        # Old schema – rename and recreate
        c.execute("ALTER TABLE results RENAME TO results_old")
        c.execute('''CREATE TABLE results
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT, name TEXT,
                      class TEXT, section TEXT, subject TEXT, score INTEGER, total_questions INTEGER,
                      percentage REAL, test_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                      chapter TEXT)''')
        c.execute("""INSERT INTO results (student_id, name, class, subject, score, total_questions, percentage)
                     SELECT student_id, name, class, subject, score, score, 0 FROM results_old""")
        c.execute("DROP TABLE results_old")
    elif 'section' not in res_columns:
        c.execute("ALTER TABLE results ADD COLUMN section TEXT")
    c.execute('''CREATE TABLE IF NOT EXISTS exam_control
                 (id INTEGER PRIMARY KEY CHECK (id=1), is_active INTEGER DEFAULT 0,
                  start_time TIMESTAMP, duration INTEGER DEFAULT 60,
                  negative_marking INTEGER DEFAULT 0,
                  negative_value REAL DEFAULT 0.33)''')
    c.execute("INSERT OR IGNORE INTO exam_control (id, is_active, start_time, duration, negative_marking, negative_value) VALUES (1, 0, NULL, 60, 0, 0.33)")
    # Add columns if upgrading from older schema
    c.execute("PRAGMA table_info(exam_control)")
    ec_cols = [r[1] for r in c.fetchall()]
    if 'negative_marking' not in ec_cols:
        c.execute("ALTER TABLE exam_control ADD COLUMN negative_marking INTEGER DEFAULT 0")
    if 'negative_value' not in ec_cols:
        c.execute("ALTER TABLE exam_control ADD COLUMN negative_value REAL DEFAULT 0.33")

    c.execute('''CREATE TABLE IF NOT EXISTS settings
                 (key TEXT PRIMARY KEY, value TEXT)''')
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('school_name', 'RRB Group of Schools')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('logo_path', '')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('school_address', '')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('academic_session', '')")
    # Theme settings
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_primary',    '#1a237e')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_secondary',  '#3949ab')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_accent',     '#10b981')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_student',    '#1b5e20')")
    # Extended theme color settings
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_icon_color',     '#4f46e5')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_btn_bg',          '#1a237e')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_btn_text',         '#ffffff')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_heading_color',    '#1e1b4b')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_body_text',        '#1e293b')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_muted_text',       '#64748b')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_page_bg',          '#f0f4f8')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_card_bg',          '#ffffff')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_sidebar_bg',       '#1a237e')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_link_color',       '#3949ab')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_hover_color',      '#4f46e5')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_border_color',     '#e2e8f0')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_success_color',    '#16a34a')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_warning_color',    '#f59e0b')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_danger_color',     '#dc2626')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_font_web',   'Poppins')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('theme_font_export','Arial')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('developer_credit', 'Developed by Gaurav Shukla & Team')")
    c.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('active_language',  'en')")

    # Translation cache for language switcher
    c.execute('''CREATE TABLE IF NOT EXISTS translation_cache
                 (lang_code  TEXT NOT NULL,
                  string_key TEXT NOT NULL,
                  translated TEXT NOT NULL,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  PRIMARY KEY (lang_code, string_key))''')
    c.execute('''CREATE TABLE IF NOT EXISTS translation_meta
                 (lang_code        TEXT PRIMARY KEY,
                  source_hash      TEXT,
                  status           TEXT DEFAULT 'pending',
                  progress         INTEGER DEFAULT 0,
                  total_strings    INTEGER DEFAULT 0,
                  translated_count INTEGER DEFAULT 0,
                  last_updated     TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    c.execute('''CREATE TABLE IF NOT EXISTS reattempt_requests
                (id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id TEXT NOT NULL,
                class TEXT NOT NULL,
                subject TEXT NOT NULL,
                status TEXT DEFAULT 'pending',
                requested_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                reviewed_at TIMESTAMP,
                admin_note TEXT)''')

    # Table for storing shuffled question order and option mapping per student
    c.execute('''CREATE TABLE IF NOT EXISTS shuffled_questions
                 (student_id TEXT,
                  question_id INTEGER,
                  shuffled_index INTEGER,
                  option_order TEXT,
                  PRIMARY KEY (student_id, question_id))''')

    # Teachers table
    c.execute('''CREATE TABLE IF NOT EXISTS teachers
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT NOT NULL,
                  mobile TEXT UNIQUE NOT NULL,
                  password TEXT NOT NULL,
                  email TEXT,
                  address TEXT,
                  picture TEXT,
                  status TEXT DEFAULT 'active',
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Teacher assignments (class and subject assignments)
    c.execute('''CREATE TABLE IF NOT EXISTS teacher_assignments
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_id INTEGER NOT NULL,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  subject TEXT NOT NULL,
                  is_class_teacher INTEGER DEFAULT 0,
                  FOREIGN KEY (teacher_id) REFERENCES teachers(id))''')
    c.execute("PRAGMA table_info(teacher_assignments)")
    ta_cols = [col[1] for col in c.fetchall()]
    if 'section' not in ta_cols:
        c.execute("ALTER TABLE teacher_assignments ADD COLUMN section TEXT DEFAULT ''")

    # Test papers table (CSV uploads with naming convention)
    c.execute('''CREATE TABLE IF NOT EXISTS test_papers
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  filename TEXT NOT NULL,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  subject TEXT NOT NULL,
                  test_no TEXT NOT NULL,
                  uploaded_by TEXT NOT NULL,
                  uploader_type TEXT DEFAULT 'admin',
                  is_active INTEGER DEFAULT 0,
                  question_count INTEGER DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Bulletins / Announcements (Admin + Teacher can post; shown on student login page)
    c.execute('''CREATE TABLE IF NOT EXISTS bulletins
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  title TEXT NOT NULL,
                  content TEXT NOT NULL,
                  posted_by TEXT NOT NULL,
                  poster_type TEXT DEFAULT 'admin',
                  target_class TEXT DEFAULT '',
                  is_active INTEGER DEFAULT 1,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Test Generation History (for AI-generated tests)
    c.execute('''CREATE TABLE IF NOT EXISTS test_generation_history
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_id TEXT NOT NULL,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  subject TEXT NOT NULL,
                  chapter TEXT NOT NULL,
                  test_no TEXT NOT NULL,
                  output_mode TEXT DEFAULT 'cbt',
                  total_questions INTEGER DEFAULT 0,
                  mcq_count INTEGER DEFAULT 0,
                  assertion_count INTEGER DEFAULT 0,
                  very_short_count INTEGER DEFAULT 0,
                  short_count INTEGER DEFAULT 0,
                  long_count INTEGER DEFAULT 0,
                  case_study_count INTEGER DEFAULT 0,
                  remark TEXT DEFAULT '',
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Ensure paper_type exists in test_papers and test_generation_history
    c.execute("PRAGMA table_info(test_papers)")
    tp_cols = [col[1] for col in c.fetchall()]
    if 'paper_type' not in tp_cols:
        c.execute("ALTER TABLE test_papers ADD COLUMN paper_type TEXT DEFAULT 'Descriptive Paper'")

    c.execute("PRAGMA table_info(test_generation_history)")
    tgh_cols = [col[1] for col in c.fetchall()]
    if 'paper_type' not in tgh_cols:
        c.execute("ALTER TABLE test_generation_history ADD COLUMN paper_type TEXT DEFAULT 'Descriptive Paper'")

    # MCQ Test History (for tracking generated MCQ tests)
    c.execute('''CREATE TABLE IF NOT EXISTS mcq_test_history
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_id INTEGER,
                  teacher_name TEXT,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  subject TEXT NOT NULL,
                  test_no TEXT DEFAULT '',
                  question_count INTEGER DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Teacher Notifications table (for background job completion alerts)
    c.execute('''CREATE TABLE IF NOT EXISTS teacher_notifications
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_id INTEGER NOT NULL,
                  title TEXT NOT NULL,
                  message TEXT NOT NULL,
                  link TEXT DEFAULT '',
                  is_read INTEGER DEFAULT 0,
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Audit Logs table (for security and action monitoring)
    c.execute('''CREATE TABLE IF NOT EXISTS audit_logs
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  user_type TEXT NOT NULL,
                  user_id TEXT NOT NULL,
                  action TEXT NOT NULL,
                  target_table TEXT DEFAULT '',
                  target_id TEXT DEFAULT '',
                  timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  ip_address TEXT DEFAULT '')''')

    # Add question_type column to questions if not exists
    c.execute("PRAGMA table_info(questions)")
    q_cols = [col[1] for col in c.fetchall()]
    if 'question_type' not in q_cols:
        c.execute("ALTER TABLE questions ADD COLUMN question_type TEXT DEFAULT 'MCQ'")
    if 'test_no' not in q_cols:
        c.execute("ALTER TABLE questions ADD COLUMN test_no TEXT DEFAULT ''")
    if 'section' not in q_cols:
        c.execute("ALTER TABLE questions ADD COLUMN section TEXT DEFAULT ''")

    # Student class lock: once registered, class/section is locked
    c.execute('''CREATE TABLE IF NOT EXISTS student_class_lock
                 (student_id TEXT PRIMARY KEY,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  locked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # Combined / Multi-Subject Test entities
    c.execute('''CREATE TABLE IF NOT EXISTS combined_tests
                 (id           INTEGER PRIMARY KEY AUTOINCREMENT,
                  test_no      TEXT NOT NULL,
                  class        TEXT NOT NULL,
                  section      TEXT DEFAULT '',
                  title        TEXT DEFAULT '',
                  created_by   TEXT NOT NULL,
                  creator_type TEXT DEFAULT 'teacher',
                  is_active    INTEGER DEFAULT 1,
                  created_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    c.execute('''CREATE TABLE IF NOT EXISTS combined_test_subjects
                 (id               INTEGER PRIMARY KEY AUTOINCREMENT,
                  combined_test_id INTEGER NOT NULL,
                  subject          TEXT NOT NULL,
                  question_count   INTEGER DEFAULT 0,
                  FOREIGN KEY (combined_test_id) REFERENCES combined_tests(id) ON DELETE CASCADE)''')

    # Scheduled Tests table (for scheduling test windows & duration control)
    c.execute('''CREATE TABLE IF NOT EXISTS scheduled_tests
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  subject TEXT NOT NULL,
                  test_no TEXT NOT NULL,
                  scheduled_date TEXT NOT NULL,
                  start_time TEXT NOT NULL,
                  end_time TEXT NOT NULL,
                  duration_minutes INTEGER NOT NULL,
                  status TEXT DEFAULT 'scheduled',
                  created_by TEXT DEFAULT 'admin',
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # ── Per-Session Exam Control ────────────────────────────────────────
    # Each teacher/admin start_exam call creates ONE row here, allowing
    # multiple independent sessions to run simultaneously on the same server.
    c.execute('''CREATE TABLE IF NOT EXISTS exam_sessions
                 (id                INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_id        INTEGER,
                  started_by_type   TEXT DEFAULT 'teacher',
                  started_by_name   TEXT,
                  lab_ip            TEXT NOT NULL,
                  lab_label         TEXT,
                  class             TEXT NOT NULL,
                  section           TEXT DEFAULT '',
                  subject           TEXT NOT NULL,
                  test_no           TEXT DEFAULT '',
                  allotted_duration INTEGER NOT NULL,
                  start_time        TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                  pause_started_at  TIMESTAMP,
                  total_paused_secs INTEGER DEFAULT 0,
                  status            TEXT DEFAULT 'running',
                  end_time          TIMESTAMP,
                  end_reason        TEXT,
                  end_message       TEXT,
                  negative_marking  INTEGER DEFAULT 0,
                  negative_value    REAL DEFAULT 0.33,
                  created_at        TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    # ── Coding Test Module Schema ──────────────────────────────────────
    c.execute('''CREATE TABLE IF NOT EXISTS coding_questions
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  test_no TEXT NOT NULL,
                  class TEXT NOT NULL,
                  section TEXT DEFAULT '',
                  subject TEXT NOT NULL,
                  title TEXT NOT NULL,
                  problem_statement TEXT NOT NULL,
                  allowed_languages TEXT DEFAULT 'python,java,c,cpp',
                  sample_input TEXT DEFAULT '',
                  sample_output TEXT DEFAULT '',
                  total_marks INTEGER DEFAULT 10,
                  created_by TEXT DEFAULT 'admin',
                  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

    c.execute("PRAGMA table_info(coding_questions)")
    cq_cols = [col[1] for col in c.fetchall()]
    for col_name, col_def in [
        ('topic', "TEXT DEFAULT ''"),
        ('difficulty', "TEXT DEFAULT 'medium'"),
        ('input_format', "TEXT DEFAULT ''"),
        ('output_format', "TEXT DEFAULT ''"),
        ('constraints', "TEXT DEFAULT ''"),
        ('ai_generated', "INTEGER DEFAULT 0"),
        ('approved', "INTEGER DEFAULT 0")
    ]:
        if col_name not in cq_cols:
            c.execute(f"ALTER TABLE coding_questions ADD COLUMN {col_name} {col_def}")

    c.execute('''CREATE TABLE IF NOT EXISTS coding_test_cases
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  question_id INTEGER NOT NULL,
                  input TEXT DEFAULT '',
                  expected_output TEXT DEFAULT '',
                  is_hidden INTEGER DEFAULT 0,
                  marks_weight REAL DEFAULT 1.0,
                  FOREIGN KEY(question_id) REFERENCES coding_questions(id) ON DELETE CASCADE)''')

    c.execute('''CREATE TABLE IF NOT EXISTS coding_submissions
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  student_id TEXT NOT NULL,
                  question_id INTEGER NOT NULL,
                  code TEXT NOT NULL,
                  language TEXT NOT NULL,
                  test_cases_passed INTEGER DEFAULT 0,
                  total_test_cases INTEGER DEFAULT 0,
                  marks_awarded REAL DEFAULT 0.0,
                  status TEXT DEFAULT 'Pending',
                  stdout_output TEXT DEFAULT '',
                  stderr_output TEXT DEFAULT '',
                  submitted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  FOREIGN KEY(question_id) REFERENCES coding_questions(id) ON DELETE CASCADE)''')

    c.execute('''CREATE TABLE IF NOT EXISTS coding_drafts
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  student_id TEXT NOT NULL,
                  question_id INTEGER NOT NULL,
                  language TEXT NOT NULL,
                  code TEXT NOT NULL,
                  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  UNIQUE(student_id, question_id, language))''')

    # Ensure board column exists in questions, test_generation_history, and test_papers
    for tbl in ('questions', 'test_generation_history', 'test_papers'):
        try:
            c.execute(f"ALTER TABLE {tbl} ADD COLUMN board TEXT DEFAULT 'CBSE'")
        except Exception:
            pass

    # ── Question Bank Module Tables ───────────────────────────────────────
    c.execute('''CREATE TABLE IF NOT EXISTS question_bank_papers
                 (id               INTEGER PRIMARY KEY AUTOINCREMENT,
                  uploaded_by      TEXT NOT NULL,
                  uploader_type    TEXT DEFAULT 'teacher',
                  original_filename TEXT NOT NULL,
                  board            TEXT NOT NULL,
                  class            TEXT NOT NULL,
                  subject          TEXT NOT NULL,
                  chapter          TEXT NOT NULL,
                  question_type    TEXT NOT NULL,
                  upload_date      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  status           TEXT DEFAULT 'pending')''')

    c.execute("""CREATE TABLE IF NOT EXISTS question_bank_items
                 (id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                  paper_id           INTEGER,
                  board              TEXT NOT NULL,
                  class              TEXT NOT NULL,
                  subject            TEXT NOT NULL,
                  chapter            TEXT NOT NULL,
                  question_type      TEXT NOT NULL,
                  raw_question_text  TEXT NOT NULL,
                  raw_options        TEXT DEFAULT '[]',
                  raw_answer         TEXT DEFAULT '',
                  image_path         TEXT DEFAULT '',
                  ai_generated       INTEGER DEFAULT 0,
                  approved           INTEGER DEFAULT 0,
                  approved_by        TEXT DEFAULT '',
                  approved_at        TIMESTAMP,
                  source_page        INTEGER DEFAULT 0,
                  created_at         TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  FOREIGN KEY (paper_id) REFERENCES question_bank_papers(id) ON DELETE SET NULL)""")

    # ── Multi-Section Test Tables & Migrations ──────────────────────────────
    c.execute("""CREATE TABLE IF NOT EXISTS test_sections
                 (id               INTEGER PRIMARY KEY AUTOINCREMENT,
                  test_paper_id    INTEGER NOT NULL,
                  section_name     TEXT NOT NULL,
                  subject          TEXT NOT NULL,
                  question_count   INTEGER NOT NULL,
                  positive_marks   REAL DEFAULT 1.0,
                  negative_marks   REAL DEFAULT 0.0,
                  duration_minutes INTEGER DEFAULT 0,
                  order_index      INTEGER DEFAULT 0,
                  generation_mode  TEXT DEFAULT 'ai',
                  FOREIGN KEY (test_paper_id) REFERENCES test_papers(id) ON DELETE CASCADE)""")

    c.execute("""CREATE TABLE IF NOT EXISTS test_section_config
                 (id                     INTEGER PRIMARY KEY AUTOINCREMENT,
                  test_paper_id          INTEGER NOT NULL UNIQUE,
                  timing_mode            TEXT DEFAULT 'disabled',
                  shuffle_sections       INTEGER DEFAULT 0,
                  two_stage_instructions INTEGER DEFAULT 1,
                  FOREIGN KEY (test_paper_id) REFERENCES test_papers(id) ON DELETE CASCADE)""")

    c.execute("""CREATE TABLE IF NOT EXISTS student_section_progress
                 (id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                  student_id         TEXT NOT NULL,
                  test_paper_id      INTEGER NOT NULL,
                  current_section_id INTEGER,
                  locked_sections    TEXT DEFAULT '[]',
                  section_time_left  TEXT DEFAULT '{}',
                  updated_at         TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                  UNIQUE(student_id, test_paper_id))""")

    # Column migrations
    c.execute("PRAGMA table_info(results)")
    res_cols = [r[1] for r in c.fetchall()]
    if 'section_breakdown' not in res_cols:
        c.execute("ALTER TABLE results ADD COLUMN section_breakdown TEXT DEFAULT ''")

    c.execute("PRAGMA table_info(students)")
    stu_cols = [r[1] for r in c.fetchall()]
    if 'current_section' not in stu_cols:
        c.execute("ALTER TABLE students ADD COLUMN current_section TEXT DEFAULT ''")

    c.execute("PRAGMA table_info(questions)")
    q_cols = [r[1] for r in c.fetchall()]
    if 'section_id' not in q_cols:
        c.execute("ALTER TABLE questions ADD COLUMN section_id INTEGER DEFAULT NULL")

    conn.commit()
    conn.close()

def update_scheduled_tests_status():
    """
    Evaluates and updates scheduled_tests status based on current datetime (scheduled -> active -> expired).
    """
    try:
        now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = get_db()
        c = conn.cursor()
        c.execute("""UPDATE scheduled_tests
                     SET status='active'
                     WHERE status='scheduled' AND start_time <= ? AND end_time >= ?""", (now_str, now_str))
        c.execute("""UPDATE scheduled_tests
                     SET status='expired'
                     WHERE status IN ('scheduled', 'active') AND end_time < ?""", (now_str,))
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.error(f"Error updating scheduled tests status: {e}")

def get_student_schedule_status(class_, section, subject, test_no, student_id=None):
    """
    Check if a test schedule exists for given class, section, subject, test_no.
    Returns dict with state ('none', 'upcoming', 'active', 'expired'), schedule info, and effective duration.
    """
    update_scheduled_tests_status()
    conn = get_db()
    c = conn.cursor()
    clean_cls = str(class_ or '').strip()
    sec = str(section or '').strip()
    subj = str(subject or '').strip()
    tno = str(test_no or '').strip()

    c.execute("""SELECT * FROM scheduled_tests
                 WHERE class=? AND (? = '' OR section=? OR section='') 
                   AND LOWER(subject)=LOWER(?) AND test_no=?
                   AND status != 'cancelled'
                 ORDER BY id DESC LIMIT 1""", (clean_cls, sec, sec, subj, tno))
    row = c.fetchone()
    conn.close()

    if not row:
        return {'state': 'none', 'schedule': None}

    sch = dict(row)
    now = datetime.datetime.now()

    try:
        if len(sch['start_time']) == 5:
            st_dt = datetime.datetime.strptime(f"{sch['scheduled_date']} {sch['start_time']}", '%Y-%m-%d %H:%M')
        else:
            st_dt = datetime.datetime.strptime(sch['start_time'][:19], '%Y-%m-%d %H:%M:%S')
    except Exception:
        st_dt = now

    try:
        if len(sch['end_time']) == 5:
            end_dt = datetime.datetime.strptime(f"{sch['scheduled_date']} {sch['end_time']}", '%Y-%m-%d %H:%M')
        else:
            end_dt = datetime.datetime.strptime(sch['end_time'][:19], '%Y-%m-%d %H:%M:%S')
    except Exception:
        end_dt = now

    sch['start_dt_str'] = st_dt.strftime('%d/%m/%Y %I:%M %p')
    sch['end_dt_str'] = end_dt.strftime('%d/%m/%Y %I:%M %p')

    if now < st_dt:
        return {
            'state': 'upcoming',
            'schedule': sch,
            'message': f"This test is locked. Scheduled for {st_dt.strftime('%d %b %Y')} at {st_dt.strftime('%I:%M %p')}."
        }
    elif st_dt <= now <= end_dt:
        rem_seconds = (end_dt - now).total_seconds()
        rem_minutes = max(1, int(rem_seconds // 60))
        effective_duration = min(sch['duration_minutes'], rem_minutes)
        return {
            'state': 'active',
            'schedule': sch,
            'effective_duration': effective_duration,
            'message': f"Available Now (Window closes at {end_dt.strftime('%I:%M %p')})"
        }
    else:
        return {
            'state': 'expired',
            'schedule': sch,
            'message': "This test window has closed and is now permanently locked. Contact admin for rescheduling."
        }

def get_db():
    """
    Get a database connection with Row factory enabled for dict-like access.
    Configured with timeout and busy_timeout to prevent lock conflicts.
    """
    conn = sqlite3.connect('database.db', timeout=30.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn

def reset_exam_on_startup():
    """
    Reset the exam status on application startup to ensure clean state.
    Also terminates any orphaned running/paused sessions from previous server runs.
    """
def reset_exam_on_startup():
    """Do not forcibly terminate active exam sessions on application reloads."""
    pass

init_db()

# ── Per-Session Exam Control Helpers ─────────────────────────────────────
def parse_session_dt(val):
    """Safely parse SQLite timestamp strings or datetime objects."""
    if isinstance(val, datetime.datetime):
        return val
    if not val:
        return datetime.datetime.now()
    s = str(val).strip().replace('T', ' ')
    for fmt in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            return datetime.datetime.strptime(s[:19] if fmt != '%Y-%m-%d %H:%M:%S.%f' else s, fmt)
        except Exception:
            pass
    return datetime.datetime.now()

def get_session_remaining_seconds(sess):
    """
    Calculate remaining time in seconds for an exam session, accounting for
    allotted_duration, start_time, pause_started_at, and total_paused_secs.
    """
    if not sess:
        return 0
    if sess.get('status') == 'ended':
        return 0
    now = datetime.datetime.now()
    st = parse_session_dt(sess.get('start_time'))
    total_paused = int(sess.get('total_paused_secs') or 0)
    allotted_secs = int(sess.get('allotted_duration') or 60) * 60

    if sess.get('status') == 'paused':
        pause_st = parse_session_dt(sess.get('pause_started_at')) if sess.get('pause_started_at') else now
        elapsed = (pause_st - st).total_seconds() - total_paused
    else:  # 'running'
        elapsed = (now - st).total_seconds() - total_paused

    rem = allotted_secs - max(0, elapsed)
    return max(0, int(rem))

def get_active_session_for_student(student_id=None, class_=None, section=None, subject=None, test_no=None, conn=None):
    """
    Find the most recent running or paused exam session matching the student's scope.
    """
    close_conn = False
    if conn is None:
        conn = get_db()
        close_conn = True
    try:
        c = conn.cursor()
        if student_id and not (class_ and subject):
            c.execute("SELECT class, section, subject, test_no FROM students WHERE student_id=?", (student_id,))
            st = c.fetchone()
            if st:
                class_ = class_ or st['class']
                section = section if section is not None else st['section']
                subject = subject or st['subject']
                test_no = test_no if test_no is not None else (st['test_no'] if 'test_no' in st.keys() else '')

        if not (class_ and subject):
            return None

        clean_cls = str(class_).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        sec = str(section or '').strip()
        subj = str(subject or '').strip()
        tno = str(test_no or '').strip()

        query = """
            SELECT * FROM exam_sessions
            WHERE status IN ('running', 'paused')
              AND (class = ? OR class = ? OR ? LIKE '%' || class || '%')
              AND (section = '' OR section IS NULL OR LOWER(section) = LOWER(?))
              AND LOWER(subject) = LOWER(?)
              AND (test_no = '' OR test_no IS NULL OR test_no = ?)
            ORDER BY id DESC LIMIT 1
        """
        c.execute(query, (class_, clean_cls, class_, sec, subj, tno))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        if close_conn:
            conn.close()

def get_latest_session_for_student(student_id=None, class_=None, section=None, subject=None, test_no=None, conn=None):
    """
    Find the most recent session for the student within the last 24h regardless of status.
    Useful for checking if a session was ended by admin/teacher with an end_message.
    """
    close_conn = False
    if conn is None:
        conn = get_db()
        close_conn = True
    try:
        c = conn.cursor()
        if student_id and not (class_ and subject):
            c.execute("SELECT class, section, subject, test_no FROM students WHERE student_id=?", (student_id,))
            st = c.fetchone()
            if st:
                class_ = class_ or st['class']
                section = section if section is not None else st['section']
                subject = subject or st['subject']
                test_no = test_no if test_no is not None else (st['test_no'] if 'test_no' in st.keys() else '')

        if not (class_ and subject):
            return None

        clean_cls = str(class_).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        sec = str(section or '').strip()
        subj = str(subject or '').strip()
        tno = str(test_no or '').strip()

        query = """
            SELECT * FROM exam_sessions
            WHERE (class = ? OR class = ? OR ? LIKE '%' || class || '%')
              AND (section = '' OR section IS NULL OR LOWER(section) = LOWER(?))
              AND LOWER(subject) = LOWER(?)
              AND (test_no = '' OR test_no IS NULL OR test_no = ?)
              AND created_at >= datetime('now', '-24 hours')
            ORDER BY id DESC LIMIT 1
        """
        c.execute(query, (class_, clean_cls, class_, sec, subj, tno))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        if close_conn:
            conn.close()

def get_session_student_counts(sess, conn=None):
    """
    Returns counts of In Progress, Submitted, and Total students matching session scope.
    """
    close_conn = False
    if conn is None:
        conn = get_db()
        close_conn = True
    try:
        c = conn.cursor()
        cls = sess.get('class', '')
        clean_cls = str(cls).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        sec = sess.get('section', '')
        subj = sess.get('subject', '')
        tno = sess.get('test_no', '')

        # Base filter
        where = "(class = ? OR class = ? OR ? LIKE '%' || class || '%') AND LOWER(subject) = LOWER(?)"
        params = [cls, clean_cls, cls, subj]
        if sec:
            where += " AND LOWER(section) = LOWER(?)"
            params.append(sec)
        if tno:
            where += " AND (test_no = ? OR test_no = '' OR test_no IS NULL)"
            params.append(tno)

        c.execute(f"SELECT COUNT(*) FROM students WHERE status='In Progress' AND {where}", params)
        in_prog = c.fetchone()[0]

        c.execute(f"SELECT COUNT(*) FROM students WHERE status='Submitted' AND {where}", params)
        submitted = c.fetchone()[0]

        c.execute(f"SELECT COUNT(*) FROM students WHERE {where}", params)
        total = c.fetchone()[0]

        return {'in_progress': in_prog, 'submitted': submitted, 'total': total}
    finally:
        if close_conn:
            conn.close()


def get_setting(key, default=''):
    """
    Retrieve a setting value from the settings table by key.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = c.fetchone()
    conn.close()
    return row['value'] if row else default

def set_setting(key, value):
    """
    Save or update a setting in the settings table.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

# ---------------------------------------------------------------------------
# SITE CONFIG — single source of truth for all white-label values
# Callable from both Flask routes (via context processor) AND export code
# ---------------------------------------------------------------------------
_GOOGLE_FONTS_MAP = {
    'Poppins':           'Poppins:wght@400;500;600;700',
    'Plus Jakarta Sans': 'Plus+Jakarta+Sans:wght@400;500;600;700;800',
    'Inter':             'Inter:wght@300;400;500;600;700',
    'Roboto':            'Roboto:wght@300;400;500;700',
    'Nunito':            'Nunito:wght@400;600;700;800',
}
_FALLBACK_CREDIT = 'Developed by Gaurav Shukla &amp; Team'

def get_site_config() -> dict:
    """
    Returns all white-label config values with safe defaults.
    Callable from Python routes, export functions, and context processor.
    """
    return {
        # Branding
        'school_name':       get_setting('school_name',       'RRB Group of Schools'),
        'school_address':    get_setting('school_address',    ''),
        'academic_session':  get_setting('academic_session',  ''),
        'logo_path':         get_setting('logo_path',         ''),
        # Theme colors - Brand
        'theme_primary':     get_setting('theme_primary',     '#1a237e'),
        'theme_secondary':   get_setting('theme_secondary',   '#3949ab'),
        'theme_accent':      get_setting('theme_accent',      '#10b981'),
        'theme_student':     get_setting('theme_student',     '#1b5e20'),
        # Theme colors - Extended
        'theme_icon_color':  get_setting('theme_icon_color',  '#4f46e5'),
        'theme_btn_bg':      get_setting('theme_btn_bg',      '#1a237e'),
        'theme_btn_text':    get_setting('theme_btn_text',    '#ffffff'),
        'theme_heading_color':   get_setting('theme_heading_color',  '#1e1b4b'),
        'theme_body_text':       get_setting('theme_body_text',      '#1e293b'),
        'theme_muted_text':      get_setting('theme_muted_text',     '#64748b'),
        'theme_page_bg':         get_setting('theme_page_bg',        '#f0f4f8'),
        'theme_card_bg':         get_setting('theme_card_bg',        '#ffffff'),
        'theme_sidebar_bg':      get_setting('theme_sidebar_bg',     '#1a237e'),
        'theme_link_color':      get_setting('theme_link_color',     '#3949ab'),
        'theme_hover_color':     get_setting('theme_hover_color',    '#4f46e5'),
        'theme_border_color':    get_setting('theme_border_color',   '#e2e8f0'),
        'theme_success_color':   get_setting('theme_success_color',  '#16a34a'),
        'theme_warning_color':   get_setting('theme_warning_color',  '#f59e0b'),
        'theme_danger_color':    get_setting('theme_danger_color',   '#dc2626'),
        # Fonts
        'theme_font_web':    get_setting('theme_font_web',    'Poppins'),
        'theme_font_export': get_setting('theme_font_export', 'Arial'),
        # Developer credit (HTML-safe version for direct template injection)
        'developer_credit':  get_setting('developer_credit',  'Developed by Gaurav Shukla &amp; Team'),
        # Language
        'active_language':   get_setting('active_language',   'en'),
    }


@app.context_processor
def inject_branding():
    """
    Injects all white-label + theme variables into every Jinja template.
    Also provides the `t()` translation helper and a `theme_style_tag` for CSS overrides.
    Falls back to safe defaults if the settings table is missing any key.
    """
    cfg = get_site_config()

    # ── Logo base64 ──────────────────────────────────────────────────────
    logo_path = cfg['logo_path']
    logo_base64 = None
    if logo_path:
        try:
            abs_path = os.path.join(app.static_folder, logo_path)
            if os.path.exists(abs_path):
                with open(abs_path, 'rb') as _f:
                    _lb   = _f.read()
                    _ext  = os.path.splitext(abs_path)[1].lower()
                    _mime = 'image/png' if _ext == '.png' else 'image/jpeg'
                    logo_base64 = f"data:{_mime};base64,{base64.b64encode(_lb).decode('utf-8')}"
        except Exception:
            logo_base64 = None

    # ── Google Fonts <link> tag ──────────────────────────────────────────
    font_web = cfg['theme_font_web']
    gf_param = _GOOGLE_FONTS_MAP.get(font_web)
    if gf_param:
        font_link_tag = (
            f'<link href="https://fonts.googleapis.com/css2?family={gf_param}&display=swap" '
            f'rel="stylesheet">'
        )
    else:
        font_link_tag = ''  # Custom / system font — no Google Fonts link needed

    # Compute darkened shade of primary for --primary-dark
    def _darken(hex_col, factor=0.75):
        hex_col = hex_col.lstrip('#')
        try:
            r, g, b = int(hex_col[0:2], 16), int(hex_col[2:4], 16), int(hex_col[4:6], 16)
            return '#{:02x}{:02x}{:02x}'.format(int(r*factor), int(g*factor), int(b*factor))
        except Exception:
            return '#0d1457'

    primary     = cfg['theme_primary']
    secondary   = cfg['theme_secondary']
    accent      = cfg['theme_accent']
    student_col = cfg['theme_student']
    icon_color  = cfg.get('theme_icon_color',  '#4f46e5')
    btn_bg      = cfg.get('theme_btn_bg',      '#1a237e')
    btn_text    = cfg.get('theme_btn_text',    '#ffffff')
    heading_col = cfg.get('theme_heading_color','#1e1b4b')
    body_text   = cfg.get('theme_body_text',   '#1e293b')
    muted_text  = cfg.get('theme_muted_text',  '#64748b')
    page_bg     = cfg.get('theme_page_bg',     '#f0f4f8')
    card_bg     = cfg.get('theme_card_bg',     '#ffffff')
    sidebar_bg  = cfg.get('theme_sidebar_bg',  '#1a237e')
    link_color  = cfg.get('theme_link_color',  '#3949ab')
    hover_color = cfg.get('theme_hover_color', '#4f46e5')
    border_color= cfg.get('theme_border_color','#e2e8f0')
    success_col = cfg.get('theme_success_color','#16a34a')
    warning_col = cfg.get('theme_warning_color','#f59e0b')
    danger_col  = cfg.get('theme_danger_color', '#dc2626')
    primary_dark = _darken(primary)

    # ── Injected <style> block (overrides per-template :root definitions) ─
    css_font_family = f"'{font_web}', sans-serif" if font_web in _GOOGLE_FONTS_MAP else f'{font_web}, sans-serif'
    theme_style_tag = f"""{font_link_tag}
<style id="site-theme-vars">
:root {{
  /* Brand Colors */
  --primary:        {primary};
  --primary-light:  {secondary};
  --primary-lighter:{secondary};
  --primary-dark:   {primary_dark};
  --accent:         {accent};
  --green-dk:       {student_col};
  --green:          {student_col};
  /* Extended Theme Colors */
  --theme-icon-color:    {icon_color};
  --theme-btn-bg:        {btn_bg};
  --theme-btn-text:      {btn_text};
  --theme-heading-color: {heading_col};
  --theme-body-text:     {body_text};
  --theme-muted-text:    {muted_text};
  --theme-page-bg:       {page_bg};
  --theme-card-bg:       {card_bg};
  --theme-sidebar-bg:    {sidebar_bg};
  --theme-link-color:    {link_color};
  --theme-hover-color:   {hover_color};
  --theme-border-color:  {border_color};
  --theme-success:       {success_col};
  --theme-warning:       {warning_col};
  --theme-danger:        {danger_col};
  --theme-font:     {css_font_family};
}}
body {{ font-family: {css_font_family}; background: {page_bg}; color: {body_text}; }}
</style>"""

    # ── t() translation helper ───────────────────────────────────────────
    active_lang = session.get('student_lang') or cfg['active_language']

    def t(key, fallback=None):
        """Return translated string for given key, falling back to English."""
        if active_lang != 'en':
            try:
                conn = get_db()
                cur  = conn.cursor()
                cur.execute(
                    'SELECT translated FROM translation_cache WHERE lang_code=? AND string_key=?',
                    (active_lang, key)
                )
                row = cur.fetchone()
                conn.close()
                if row and row[0]:
                    return row[0]
            except Exception:
                pass
        # Fallback: English
        try:
            from ui_strings import UI_STRINGS
            return UI_STRINGS.get(key, fallback or key)
        except Exception:
            return fallback or key

    return dict(
        school_name=cfg['school_name'],
        school_address=cfg['school_address'],
        academic_session=cfg['academic_session'],
        logo_path=logo_path,
        logo_base64=logo_base64,
        theme_primary=primary,
        theme_secondary=secondary,
        theme_accent=accent,
        theme_student=student_col,
        theme_icon_color=icon_color,
        theme_btn_bg=btn_bg,
        theme_btn_text=btn_text,
        theme_heading_color=heading_col,
        theme_body_text=body_text,
        theme_muted_text=muted_text,
        theme_page_bg=page_bg,
        theme_card_bg=card_bg,
        theme_sidebar_bg=sidebar_bg,
        theme_link_color=link_color,
        theme_hover_color=hover_color,
        theme_border_color=border_color,
        theme_success_color=success_col,
        theme_warning_color=warning_col,
        theme_danger_color=danger_col,
        theme_font_web=font_web,
        theme_font_export=cfg['theme_font_export'],
        developer_credit=cfg['developer_credit'],
        active_language=active_lang,
        theme_style_tag=theme_style_tag,
        t=t,
    )


def log_audit_event(user_type, user_id, action, target_table='', target_id='', ip_address=''):

    """
    Inserts record into audit_logs table for audit tracking.
    """
    try:
        if not ip_address and request:
            ip_address = request.remote_addr or ''
        conn = get_db()
        c = conn.cursor()
        c.execute("""INSERT INTO audit_logs (user_type, user_id, action, target_table, target_id, ip_address)
                     VALUES (?, ?, ?, ?, ?, ?)""",
                  (str(user_type or 'system'), str(user_id or 'anon'), str(action),
                   str(target_table or ''), str(target_id or ''), str(ip_address or '')))
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.error(f"Failed to record audit log: {e}")

def _is_api_request():
    return request.path.startswith('/api/') or request.is_json or request.headers.get('X-Requested-With') == 'XMLHttpRequest'

def admin_required(f):
    """
    Decorator to enforce admin login requirement on routes.
    Returns 401 JSON for API requests, redirects to admin_login for browser requests.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('admin_logged_in'):
            if _is_api_request():
                return jsonify({'status': 'error', 'message': 'Unauthorized. Admin login required.'}), 401
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    return decorated_function

def student_required(f):
    """
    Decorator to enforce student login requirement on routes.
    Returns 401 JSON for API requests, redirects to index for browser requests.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('student_id'):
            if _is_api_request():
                return jsonify({'status': 'error', 'message': 'Unauthorized. Student session required.'}), 401
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def teacher_required(f):
    """
    Decorator to enforce teacher login requirement on routes.
    Returns 401 JSON for API requests, redirects to teacher_login for browser requests.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('teacher_logged_in'):
            if _is_api_request():
                return jsonify({'status': 'error', 'message': 'Unauthorized. Teacher login required.'}), 401
            return redirect(url_for('teacher_login'))
        return f(*args, **kwargs)
    return decorated_function

# -------------------------------
# Student Routes
# -------------------------------
@app.route('/')
@app.route('/login')
def index():
    """
    Render the student login page with available classes, subjects, bulletins, and class teachers.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT DISTINCT class, subject FROM questions WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class, subject FROM coding_questions WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class, subject FROM test_papers WHERE class IS NOT NULL AND class != ''
                 ORDER BY class, subject""")
    rows = c.fetchall()

    class_subject_map = {}
    for row in rows:
        cls = row['class']
        sub = row['subject']
        if cls not in class_subject_map:
            class_subject_map[cls] = []
        class_subject_map[cls].append(sub)

    # Get active bulletins for the student login page bulletin board
    c.execute("""SELECT b.title, b.content, b.poster_type, b.target_class, b.created_at,
                        COALESCE(t.name, 'Admin') as poster_name
                 FROM bulletins b
                 LEFT JOIN teachers t ON b.posted_by = t.id AND b.poster_type = 'teacher'
                 WHERE b.is_active = 1
                 ORDER BY b.created_at DESC LIMIT 10""")
    bulletins = [dict(r) for r in c.fetchall()]

    # Get class teachers for display
    c.execute("""SELECT ta.class, ta.section, t.name, t.picture
                 FROM teacher_assignments ta
                 JOIN teachers t ON ta.teacher_id = t.id
                 WHERE ta.is_class_teacher = 1
                 ORDER BY ta.class, ta.section""")
    class_teachers = [dict(r) for r in c.fetchall()]

    # Get active & upcoming test schedules to flash on Notice Board
    update_scheduled_tests_status()
    c.execute("""SELECT * FROM scheduled_tests 
                 WHERE status IN ('scheduled', 'active') 
                 ORDER BY scheduled_date ASC, start_time ASC LIMIT 10""")
    scheduled_tests = [dict(r) for r in c.fetchall()]

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')
    conn.close()

    return render_template('index.html',
                           class_subject_map=class_subject_map,
                           bulletins=bulletins,
                           scheduled_tests=scheduled_tests,
                           class_teachers=class_teachers,
                           school_name=school_name,
                           logo_path=logo_path)

# ---------------------------------------------------------------------------
# Compiled once at module level — matches option texts that reference other
# options by letter/position (e.g. "Both a and b", "None of these",
# "All of the above", "Except a and b").  Any option whose text matches
# this pattern MUST NOT be shuffled: if it moved to a different display
# slot the options it refers to would be at wrong positions, making the
# text semantically wrong (e.g. option (a) saying "Both a and b" when
# the real options it refers to are now at (c) and (d)).
_REFERENTIAL_OPT_RE = re.compile(
    r'\b(both|all\s+of\s+(the|these)|none\s+of\s+(the|these)|except|only)\b'
    r'|'
    r'\b[a-d]\b|\([a-d]\)',
    re.IGNORECASE
)

def _shuffle_options_safe(qid, cursor):
    """
    Returns a 4-char option_order string (a permutation of 'ABCD') for the
    question with the given id.

    Referential options — those whose text references other options by
    letter/position (e.g. "Both a and b", "None of these") — are kept at
    their ORIGINAL position so their meaning stays valid after the shuffle.
    Only non-referential options are shuffled among the remaining slots.
    """
    cursor.execute(
        "SELECT option_a, option_b, option_c, option_d FROM questions WHERE id=?",
        (qid,)
    )
    row = cursor.fetchone()
    if not row:
        return 'ABCD'  # Fallback: identity order if question not found

    letters = ['A', 'B', 'C', 'D']
    opt_texts = {
        'A': row['option_a'] or '',
        'B': row['option_b'] or '',
        'C': row['option_c'] or '',
        'D': row['option_d'] or '',
    }

    # Classify each slot: pinned (referential) stays, free gets shuffled.
    pinned = {letter for letter, text in opt_texts.items()
              if _REFERENTIAL_OPT_RE.search(text)}

    free_letters  = [l for l in letters if l not in pinned]
    free_shuffled = free_letters[:]
    random.shuffle(free_shuffled)

    # Rebuild full order: pinned slots keep original letter; free slots
    # receive shuffled letters in order.
    result = []
    free_iter = iter(free_shuffled)
    for letter in letters:
        if letter in pinned:
            result.append(letter)
        else:
            result.append(next(free_iter))
    return ''.join(result)
# ---------------------------------------------------------------------------

@app.route('/student_login', methods=['POST'])
@validate_schema(STUDENT_LOGIN_SCHEMA)
def student_login():
    """
    Handle student login and exam registration. Validates class lock, manages reattempt requests.
    """
    auto_submit_expired_exams()

    name        = request.form.get('name', '').strip()
    student_id  = request.form.get('student_id', '').strip()
    class_      = request.form.get('class', '').strip()
    section     = request.form.get('section', '').strip()
    subject     = request.form.get('subject', '').strip()
    test_no     = request.form.get('test_no', '').strip()   # Feature #4
    ip          = request.remote_addr

    if not all([name, student_id, class_, section, subject]):
        return "All fields required (Name, Admission No, Class, Section, Subject)", 400

    conn = get_db()
    c = conn.cursor()

    # Auto-select active or scheduled test if no specific test_no was selected by student
    if not test_no:
        clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        c.execute("""SELECT test_no FROM scheduled_tests
                     WHERE (class=? OR class LIKE ?) AND (? = '' OR section=? OR section='' OR section IS NULL)
                       AND LOWER(subject)=LOWER(?) AND status != 'cancelled'
                     ORDER BY CASE WHEN status='active' THEN 0 WHEN status='scheduled' THEN 1 ELSE 2 END, id DESC
                     LIMIT 1""",
                  (class_, f"%{clean_cls}%", section, section, subject))
        sch_row = c.fetchone()
        if sch_row and sch_row['test_no']:
            test_no = sch_row['test_no']
        else:
            conn.close()
            return render_template('index.html',
                                   error=f"No test is currently scheduled for Class {class_} ({section}) - {subject}. Students can only take scheduled tests.",
                                   class_subject_map=_get_class_subject_map(),
                                   bulletins=[], scheduled_tests=[], class_teachers=[],
                                   school_name=get_setting('school_name','RRB Group of Schools'),
                                   logo_path=get_setting('logo_path',''))

    # ── INACTIVE TEST PAPER CHECK ─────────────────────────────────
    if class_ and subject and test_no:
        clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        inact_query = """SELECT id FROM test_papers 
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? AND is_active=0"""
        inact_params = [class_, f"%{clean_cls}%", subject, test_no]
        if section:
            inact_query += " AND (section=? OR section='' OR section IS NULL)"
            inact_params.append(section)
        c.execute(inact_query, inact_params)
        if c.fetchone():
            conn.close()
            return render_template('index.html',
                                   error=f"Test '{test_no}' is currently inactive and not available. Please contact your teacher or administrator.",
                                   class_subject_map=_get_class_subject_map(),
                                   bulletins=[], scheduled_tests=[], class_teachers=[],
                                   school_name=get_setting('school_name','RRB Group of Schools'),
                                   logo_path=get_setting('logo_path',''))

    # ── SCHEDULE EXISTENCE & LOCK CHECK ───────────────────────────
    if class_ and subject and test_no:
        ms_info = _get_multi_section_info(class_, subject, test_no, conn)
        sch_res = get_student_schedule_status(class_, section, subject, test_no, student_id)
        if sch_res['state'] in ('upcoming', 'expired'):
            conn.close()
            return render_template('index.html',
                                   error=f"Test '{test_no}' is locked. {sch_res['message']}",
                                   class_subject_map=_get_class_subject_map(),
                                   bulletins=[], scheduled_tests=[], class_teachers=[],
                                   school_name=get_setting('school_name','RRB Group of Schools'),
                                   logo_path=get_setting('logo_path',''))
        elif sch_res['state'] == 'none' and not ms_info:
            conn.close()
            return render_template('index.html',
                                   error=f"Test '{test_no}' is not scheduled. Students can only take scheduled tests.",
                                   class_subject_map=_get_class_subject_map(),
                                   bulletins=[], scheduled_tests=[], class_teachers=[],
                                   school_name=get_setting('school_name','RRB Group of Schools'),
                                   logo_path=get_setting('logo_path',''))

    # ── CLASS LOCK CHECK ──────────────────────────────────────────
    c.execute("SELECT class, section FROM student_class_lock WHERE student_id=?", (student_id,))
    lock = c.fetchone()
    if lock:
        if lock['class'] != class_ or lock['section'] != section:
            conn.close()
            return render_template('index.html',
                                   error=f"You are locked to Class {lock['class']} Section {lock['section']}. Contact Admin to change.",
                                   class_subject_map=_get_class_subject_map(),
                                   bulletins=[], scheduled_tests=[], class_teachers=[],
                                   school_name=get_setting('school_name','RRB Group of Schools'),
                                   logo_path=get_setting('logo_path',''))
    else:
        c.execute("INSERT OR IGNORE INTO student_class_lock (student_id, class, section) VALUES (?,?,?)",
                  (student_id, class_, section))

    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()

    # BUG-001 FIX: Only block if student submitted THIS specific test (same subject+test_no)
    # Different test_no = new exam, must never be blocked by a previous test's submission
    if student and student['status'] == 'Submitted':
        prev_subject = student['subject'] or ''
        # Check if previous submission was for the SAME test
        c.execute("""SELECT id FROM results
                     WHERE student_id=? AND subject=? AND (
                         (? != '' AND chapter=?) OR
                         (? = '' AND subject=?)
                     ) LIMIT 1""",
                  (student_id, subject,
                   test_no, test_no,
                   test_no, subject))
        same_test_result = c.fetchone()

        if same_test_result and prev_subject == subject:
            # Same test attempted again — show reattempt prompt
            conn.commit()
            conn.close()
            session.update({'student_id': student_id, 'student_name': name,
                            'class': class_, 'subject': subject, 'section': section, 'test_no': test_no})
            return redirect(url_for('submitted'))
        else:
            # Different test (new test_no or different subject) — reset status and allow
            c.execute("""UPDATE students SET status='Not Started', exam_started_at=NULL,
                          class=?, subject=?, section=?, ip=? WHERE student_id=?""",
                      (class_, subject, section, ip, student_id))
            c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
            conn.commit()

    # ── RESUME-ON-RECONNECT ─────────────────────────────────────────────────
    # If a student is 'In Progress' on the SAME test (same subject + test_no),
    # this is a reconnect after a power cut / browser crash — NOT a fresh start.
    # We must NOT delete or reshuffle their questions; their saved answers in the
    # `responses` table are intact and will be pre-filled by /get_questions.
    # The timer in /get_exam_time already deducts elapsed time from exam_started_at,
    # so the remaining time shown on reconnect is automatically correct.
    if student and student['status'] == 'In Progress':
        prev_subject = (student['subject'] or '').strip()
        prev_test_no = ((student['test_no'] if 'test_no' in student.keys() else '') or '').strip()
        if prev_subject.lower() == subject.lower() and prev_test_no == test_no:
            # Same ongoing exam — restore session and skip reshuffle
            _sn = request.headers.get('User-Agent', 'Unknown')[:255]
            c.execute("UPDATE students SET name=?, ip=?, system_name=? WHERE student_id=?",
                      (name, ip, _sn, student_id))
            # Determine test type for correct routing (coding vs MCQ)
            _clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            c.execute("""SELECT paper_type FROM test_papers
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?)
                           AND test_no=? AND paper_type='Coding Test' LIMIT 1""",
                      (class_, f"%{_clean_cls}%", subject, test_no))
            _is_coding_resume = bool(c.fetchone())
            # Detect combined test
            _comb_id, _comb_subjs = None, []
            _ms = None
            if not _is_coding_resume and test_no:
                _ms = _get_multi_section_info(class_, subject, test_no, conn)
                if not _ms:
                    c.execute("""SELECT ct.id FROM combined_tests ct
                                 WHERE ct.test_no=? AND ct.class=? AND ct.is_active=1""",
                              (test_no, class_))
                    _ct = c.fetchone()
                    if _ct:
                        _comb_id = _ct['id']
                        c.execute("""SELECT subject FROM combined_test_subjects
                                     WHERE combined_test_id=? ORDER BY id""", (_comb_id,))
                        _comb_subjs = [r['subject'] for r in c.fetchall()]
            conn.commit()
            conn.close()
            session.update({
                'student_id': student_id, 'student_name': name,
                'class': class_, 'subject': subject, 'section': section, 'test_no': test_no,
                'test_type': 'Coding Test' if _is_coding_resume else 'MCQ',
                'combined_test_id': _comb_id,
                'combined_subjects': _comb_subjs,
                'is_multi_section': bool(_ms),
                'test_paper_id': _ms['tp_id'] if _ms else None,
                'timing_mode': _ms['timing_mode'] if _ms else 'disabled',
                'resume_mode': True
            })
            return redirect(url_for('waiting'))
    # ── END RESUME-ON-RECONNECT ─────────────────────────────────────────────

    system_name = request.headers.get('User-Agent', 'Unknown')[:255]  # limit length

    if student:
        c.execute("""UPDATE students 
                     SET name=?, class=?, subject=?, section=?, ip=?, 
                         test_no=?, system_name=?
                     WHERE student_id=?""",
                  (name, class_, subject, section, ip, test_no, system_name, student_id))
    else:
        c.execute("""INSERT INTO students 
                     (student_id, name, class, subject, section, ip, test_no, system_name, status)
                     VALUES (?,?,?,?,?,?,?,?,'Not Started')""",
                  (student_id, name, class_, subject, section, ip, test_no, system_name))

    conn.commit()
    c.execute("SELECT status, subject FROM students WHERE student_id=?", (student_id,))
    updated_status = c.fetchone()
    if updated_status and updated_status['status'] == 'Submitted' and updated_status['subject'] == subject:
        conn.close()
        session.update({'student_id': student_id, 'student_name': name,
                        'class': class_, 'subject': subject, 'section': section, 'test_no': test_no})
        return redirect(url_for('submitted'))

    # ── CODING TEST DETECTION & ROUTING ────────────────────────────
    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    c.execute("""SELECT paper_type FROM test_papers 
                 WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? AND paper_type='Coding Test' LIMIT 1""",
              (class_, f"%{clean_cls}%", subject, test_no))
    is_coding = bool(c.fetchone())

    if not is_coding and test_no:
        c.execute("""SELECT id FROM coding_questions 
                     WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? LIMIT 1""",
                  (class_, f"%{clean_cls}%", subject, test_no))
        is_coding = bool(c.fetchone())

    if is_coding:
        c.execute("""SELECT id FROM coding_questions
                     WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND (test_no=? OR topic=?)
                     ORDER BY id""",
                  (class_, f"%{clean_cls}%", subject, test_no, test_no))
        cqids = [r['id'] for r in c.fetchall()]
        if not cqids:
            c.execute("""SELECT id FROM coding_questions
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?)
                         ORDER BY id""",
                      (class_, f"%{clean_cls}%", subject))
            cqids = [r['id'] for r in c.fetchall()]

        random.shuffle(cqids)
        c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
        for idx, qid in enumerate(cqids):
            c.execute("INSERT INTO shuffled_questions (student_id, question_id, shuffled_index, option_order) VALUES (?,?,?,?)",
                       (student_id, qid, idx, ''))
        conn.commit()
        conn.close()

        session.update({'student_id': student_id, 'student_name': name,
                        'class': class_, 'subject': subject, 'section': section, 'test_no': test_no,
                        'test_type': 'Coding Test',
                        'combined_test_id': None,
                        'combined_subjects': []})
        return redirect(url_for('waiting'))

    # ── MULTI-SECTION TEST HANDLING ──────────────────────────────────────────
    ms_info = _get_multi_section_info(class_, subject, test_no, conn)
    if ms_info:
        tp_id = ms_info['tp_id']
        timing_mode = ms_info['timing_mode']
        shuffle_secs = bool(ms_info['shuffle_sections'])

        c.execute("""SELECT id, section_name, subject, question_count, duration_minutes, order_index
                     FROM test_sections
                     WHERE test_paper_id=?
                     ORDER BY order_index ASC""", (tp_id,))
        sec_rows = [dict(r) for r in c.fetchall()]

        if shuffle_secs:
            random.shuffle(sec_rows)

        c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
        global_q_idx = 0
        first_sec_id = None
        first_sec_name = ''
        first_sec_dur = 0

        for s_idx, sec in enumerate(sec_rows):
            sec_display_name = sec['section_name']
            if sec.get('subject') and sec.get('subject').lower() not in sec_display_name.lower():
                sec_display_name = f"{sec['section_name']} - {sec['subject']}"
            if s_idx == 0:
                first_sec_id = sec['id']
                first_sec_name = sec_display_name
                first_sec_dur = sec['duration_minutes'] or 0

            c.execute("""SELECT id FROM questions 
                         WHERE section_id=? ORDER BY id""", (sec['id'],))
            s_qids = [r['id'] for r in c.fetchall()]
            if not s_qids:
                c.execute("""SELECT id FROM questions 
                             WHERE class=? AND subject=? AND test_no=? 
                             ORDER BY id""", (class_, sec['subject'], test_no))
                s_qids = [r['id'] for r in c.fetchall()]

            random.shuffle(s_qids)

            for qid in s_qids:
                opt_order = _shuffle_options_safe(qid, c)
                c.execute("INSERT INTO shuffled_questions (student_id, question_id, shuffled_index, option_order) VALUES (?,?,?,?)",
                          (student_id, qid, global_q_idx, opt_order))
                global_q_idx += 1

        c.execute("""INSERT OR REPLACE INTO student_section_progress
                     (id, student_id, test_paper_id, current_section_id, locked_sections, section_time_left, updated_at)
                     VALUES (
                         (SELECT id FROM student_section_progress WHERE student_id=? AND test_paper_id=?),
                         ?, ?, ?, '[]', ?, CURRENT_TIMESTAMP
                     )""",
                  (student_id, tp_id, student_id, tp_id, first_sec_id, first_sec_dur * 60))

        if first_sec_name:
            c.execute("UPDATE students SET current_section=? WHERE student_id=?", (first_sec_name, student_id))

        conn.commit()
        conn.close()

        session.update({
            'student_id': student_id, 'student_name': name,
            'class': class_, 'subject': subject, 'section': section, 'test_no': test_no,
            'test_type': 'MCQ',
            'combined_test_id': None,
            'combined_subjects': [],
            'is_multi_section': True,
            'test_paper_id': tp_id,
            'timing_mode': timing_mode
        })
        return redirect(url_for('waiting'))
    # ── END MULTI-SECTION TEST HANDLING ──────────────────────────────────────

    # Shuffle questions — filter by test_no if provided
    # For combined tests, detect and load ALL subjects
    combined_test_id = None
    combined_subjects = []

    if test_no:
        # Check if this test_no is a combined test for this class
        c.execute("""SELECT ct.id FROM combined_tests ct
                      WHERE ct.test_no=? AND ct.class=? AND ct.is_active=1""",
                   (test_no, class_))
        ct_row = c.fetchone()
        if ct_row:
            combined_test_id = ct_row['id']
            c.execute("""SELECT subject FROM combined_test_subjects
                          WHERE combined_test_id=? ORDER BY id""", (combined_test_id,))
            combined_subjects = [r['subject'] for r in c.fetchall()]

            # Load ALL questions across all subjects for this combined test
            placeholders = ','.join(['?']*len(combined_subjects))
            c.execute(f"""SELECT id FROM questions
                           WHERE class=? AND test_no=? AND subject IN ({placeholders})
                           ORDER BY subject, id""",
                       [class_, test_no] + combined_subjects)
            qids = [r['id'] for r in c.fetchall()]
        else:
            c.execute("PRAGMA table_info(questions)")
            q_cols = [col[1] for col in c.fetchall()]
            if 'chapter' in q_cols:
                c.execute("SELECT id FROM questions WHERE class=? AND subject=? AND (test_no=? OR chapter=?) ORDER BY id",
                           (class_, subject, test_no, test_no))
            else:
                c.execute("SELECT id FROM questions WHERE class=? AND subject=? AND test_no=? ORDER BY id",
                           (class_, subject, test_no))
            qids = [r['id'] for r in c.fetchall()]
    else:
        c.execute("SELECT id FROM questions WHERE class=? AND subject=? ORDER BY id", (class_, subject))
        qids = [r['id'] for r in c.fetchall()]

    random.shuffle(qids)
    c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
    for idx, qid in enumerate(qids):
        opt_order = _shuffle_options_safe(qid, c)
        c.execute("INSERT INTO shuffled_questions (student_id, question_id, shuffled_index, option_order) VALUES (?,?,?,?)",
                   (student_id, qid, idx, opt_order))
    conn.commit()
    conn.close()

    session.update({'student_id': student_id, 'student_name': name,
                    'class': class_, 'subject': subject, 'section': section, 'test_no': test_no,
                    'test_type': 'MCQ',
                    'combined_test_id': combined_test_id,
                    'combined_subjects': combined_subjects})
    return redirect(url_for('waiting'))

def _get_class_subject_map():
    """
    Helper function to get a mapping of classes to their available subjects.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT DISTINCT class, subject FROM questions WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class, subject FROM coding_questions WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class, subject FROM test_papers WHERE class IS NOT NULL AND class != ''
                 ORDER BY class, subject""")
    rows = c.fetchall()
    conn.close()
    csmap = {}
    for r in rows:
        csmap.setdefault(r['class'], []).append(r['subject'])
    return csmap

@app.route('/api/get_subjects_for_class')
def api_get_subjects_for_class():
    """Return distinct subjects available for a given class across questions, QB, test papers, and curriculum."""
    class_ = request.args.get('class', '').strip()
    board = (request.args.get('board') or 'CBSE').strip() or 'CBSE'
    if not class_:
        return jsonify({'subjects': []})
    conn = get_db()
    c = conn.cursor()
    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    c.execute("""SELECT DISTINCT subject FROM questions WHERE class=? OR class LIKE ?
                 UNION
                 SELECT DISTINCT subject FROM question_bank_items WHERE class=? OR class LIKE ?
                 UNION
                 SELECT DISTINCT subject FROM test_papers WHERE class=? OR class LIKE ?
                 ORDER BY subject""", (class_, f"%{clean_cls}%", class_, f"%{clean_cls}%", class_, f"%{clean_cls}%"))
    db_subjs = [r['subject'] for r in c.fetchall() if r['subject']]
    conn.close()

    curr_subjs = []
    try:
        curr_subjs = curriculum_data.get_subjects_for_class(clean_cls, board) or []
    except Exception:
        pass

    seen = set()
    all_subjs = []
    for s in db_subjs + curr_subjs:
        s_clean = s.strip()
        if s_clean and s_clean.lower() not in seen:
            seen.add(s_clean.lower())
            all_subjs.append(s_clean)
    all_subjs.sort()
    return jsonify({'subjects': all_subjs})

@app.route('/api/student_lookup')
@app.route('/api/student/lookup')
def api_student_lookup():
    """
    Look up a student by admission number (or student_id) for instant login autofill.
    Searches both admission_no and student_id fields.
    """
    admission_no = request.args.get('admission_no', '').strip() or request.args.get('student_id', '').strip()
    if not admission_no:
        return jsonify({'status': 'error', 'found': False, 'message': 'Admission number is required'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT student_id, admission_no, name, class, section, house, dob, parents_name
        FROM students
        WHERE LOWER(TRIM(COALESCE(admission_no, ''))) = LOWER(TRIM(?))
           OR LOWER(TRIM(COALESCE(student_id, ''))) = LOWER(TRIM(?))
        ORDER BY CASE WHEN admission_no IS NOT NULL AND admission_no != '' THEN 0 ELSE 1 END, rowid DESC
        LIMIT 1
    """, (admission_no, admission_no))
    row = c.fetchone()
    conn.close()

    if not row:
        return jsonify({'status': 'success', 'found': False, 'message': 'Student not found in registry'})

    st = dict(row)
    return jsonify({
        'status': 'success',
        'found': True,
        'student': {
            'student_id': st.get('student_id') or admission_no,
            'admission_no': st.get('admission_no') or st.get('student_id') or admission_no,
            'name': st.get('name') or '',
            'class': st.get('class') or '',
            'section': st.get('section') or '',
            'house': st.get('house') or '',
            'dob': st.get('dob') or '',
            'parents_name': st.get('parents_name') or ''
        }
    })

@app.route('/api/test_numbers')
def api_test_numbers():
    """Return distinct test_no values for a class+subject combo with schedule lock status."""
    class_   = request.args.get('class','').strip()
    section  = request.args.get('section','').strip()
    subject  = request.args.get('subject','').strip()
    for_student = request.args.get('for_student', '').strip()
    is_student = (for_student == '1') or (not session.get('admin_logged_in') and not session.get('teacher_logged_in'))

    if not class_:
        return jsonify({'test_numbers': [], 'test_details': [], 'combined_tests': []})

    conn = get_db()
    c = conn.cursor()

    if subject:
        clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        
        # Identify all test_no values that are explicitly marked inactive (is_active=0)
        inact_query = """SELECT DISTINCT test_no FROM test_papers 
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND is_active=0 AND test_no IS NOT NULL AND test_no != ''"""
        inact_params = [class_, f"%{clean_cls}%", subject]
        if section:
            inact_query += " AND (section=? OR section='' OR section IS NULL)"
            inact_params.append(section)
        c.execute(inact_query, inact_params)
        inactive_test_nos = set(r[0] for r in c.fetchall() if r[0])

        # Active test_papers
        act_query = """SELECT DISTINCT test_no FROM test_papers 
                       WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND is_active=1 AND test_no IS NOT NULL AND test_no != ''"""
        act_params = [class_, f"%{clean_cls}%", subject]
        if section:
            act_query += " AND (section=? OR section='' OR section IS NULL)"
            act_params.append(section)
        c.execute(act_query, act_params)
        active_tp_nos = set(r[0] for r in c.fetchall() if r[0])

        # Questions table test numbers
        c.execute("""SELECT DISTINCT COALESCE(NULLIF(test_no,''), chapter) as tno FROM questions 
                     WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND (test_no IS NOT NULL AND test_no != '' OR chapter IS NOT NULL AND chapter != '')""",
                  (class_, f"%{clean_cls}%", subject))
        q_nos = set(r[0] for r in c.fetchall() if r[0])

        # Coding questions table test numbers
        c.execute("""SELECT DISTINCT test_no FROM coding_questions
                     WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no IS NOT NULL AND test_no != ''""",
                  (class_, f"%{clean_cls}%", subject))
        cq_nos = set(r[0] for r in c.fetchall() if r[0])

        # Also get any test numbers defined in scheduled_tests directly
        c.execute("""SELECT DISTINCT test_no FROM scheduled_tests
                     WHERE (class=? OR class LIKE ?) AND (? = '' OR section=? OR section='' OR section IS NULL)
                       AND LOWER(subject)=LOWER(?) AND status != 'cancelled' AND test_no IS NOT NULL AND test_no != ''""",
                  (class_, f"%{clean_cls}%", section, section, subject))
        sched_nos = set(r[0] for r in c.fetchall() if r[0])

        # Combine active and subtract any inactive test papers
        valid_nos = (active_tp_nos | q_nos | cq_nos | sched_nos) - inactive_test_nos
        test_numbers = sorted(list(valid_nos))
    else:
        test_numbers = []

    # Calculate schedule lock status for each test_no
    update_scheduled_tests_status()
    test_details = []
    for tn in test_numbers:
        sch_res = get_student_schedule_status(class_, section, subject, tn)
        state = sch_res.get('state', 'none')

        # When serving student view, only show scheduled tests!
        if is_student and state == 'none':
            continue

        is_locked = False
        status_label = ""
        msg = sch_res.get('message', '')

        if state == 'upcoming':
            is_locked = True
            sch_info = sch_res.get('schedule', {})
            st_time = sch_info.get('start_dt_str', sch_info.get('start_time', ''))
            status_label = f"🔒 Locked (Scheduled for {st_time})"
        elif state == 'expired':
            is_locked = True
            status_label = "🔴 Locked (Window Expired)"
        elif state == 'active':
            status_label = "🟢 Available Now"

        test_details.append({
            'test_no': tn,
            'is_locked': is_locked,
            'state': state,
            'status_label': status_label,
            'message': msg
        })

    # For students, test_numbers list only contains the scheduled tests
    if is_student:
        test_numbers = [d['test_no'] for d in test_details]
        combined = []
    else:
        # Combined (multi-subject) tests for this class (admin/teacher view)
        c.execute("""SELECT ct.id, ct.test_no, ct.title,
                            GROUP_CONCAT(cts.subject || '(' || cts.question_count || 'Q)', ', ') as subjects_info
                     FROM combined_tests ct
                     JOIN combined_test_subjects cts ON cts.combined_test_id = ct.id
                     WHERE ct.class=? AND ct.is_active=1
                     GROUP BY ct.id ORDER BY ct.created_at DESC""", (class_,))
        combined = [dict(r) for r in c.fetchall()]

    conn.close()
    return jsonify({'test_numbers': test_numbers, 'test_details': test_details, 'combined_tests': combined})

@app.route('/api/generate_combined_test', methods=['POST'])
@limiter.limit("13 per minute")
def generate_combined_test():
    """Generate a truly unified multi-subject test — one entity, multiple subjects."""
    if not session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401

    import urllib.request, json as json_lib

    class_   = request.form.get('class','').strip()
    section  = request.form.get('section','').strip()
    test_no  = request.form.get('test_no','').strip()
    title    = request.form.get('title','').strip() or test_no
    remark   = request.form.get('remark','').strip()

    # subjects_json: [{"subject":"Math","count":10}, ...]
    try:
        subjects = json_lib.loads(request.form.get('subjects_json','[]'))
    except Exception:
        return jsonify({'status':'error','message':'Invalid subjects data'}), 400

    if not class_ or not test_no or not subjects:
        return jsonify({'status':'error','message':'Class, Test No and at least one subject required'}), 400

    gemini_key = _get_gemini_key()
    if not gemini_key:
        return jsonify({'status':'error','message':'GEMINI_API_KEY not configured'}), 500

    conn = get_db()
    c = conn.cursor()

    # Check no combined test with same test_no+class already exists
    c.execute("SELECT id FROM combined_tests WHERE test_no=? AND class=?", (test_no, class_))
    if c.fetchone():
        conn.close()
        return jsonify({'status':'error',
                        'message':f'Combined test "{test_no}" already exists for Class {class_}. Choose a different Test No.'}), 400

    creator_id   = str(session.get('teacher_id','admin'))
    creator_type = 'teacher' if session.get('teacher_logged_in') else 'admin'

    # Create the combined_tests record first
    c.execute("""INSERT INTO combined_tests (test_no, class, section, title, created_by, creator_type)
                 VALUES (?,?,?,?,?,?)""",
             (test_no, class_, section, title, creator_id, creator_type))
    combined_id = c.lastrowid
    conn.commit()

    errors       = []
    total_inserted = 0

    for subj_entry in subjects:
        subject   = subj_entry.get('subject','').strip()
        mcq_count = int(subj_entry.get('count', 0))
        if not subject or mcq_count <= 0:
            continue

        try:
            questions, gen_err = generate_cbt_questions_in_batches(
                class_, section, subject, f'{title} — {subject}', test_no,
                remark or 'Standard CBSE difficulty',
                type_counts=[('MCQ (4 options, mark correct_answer as option_a/b/c/d)', mcq_count)],
                board='CBSE', max_batch_size=10
            )
            if not questions or not isinstance(questions, list):
                app.logger.error(f"Auto generate failed for {subject}: {gen_err}")
                continue

            inserted = 0
            for q in questions:
                if q.get('question_type','MCQ') in ('MCQ','Assertion-Reason'):
                    c.execute("""INSERT INTO questions
                               (class, subject, chapter, test_no, question_type,
                                question, option_a, option_b, option_c, option_d, correct_answer)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                             (class_, subject, title, test_no,
                              q.get('question_type','MCQ'), q.get('question',''),
                              q.get('option_a',''), q.get('option_b',''),
                              q.get('option_c',''), q.get('option_d',''),
                              q.get('correct_answer','')))
                    inserted += 1

            # Record subject in combined_test_subjects
            c.execute("""INSERT INTO combined_test_subjects (combined_test_id, subject, question_count)
                         VALUES (?,?,?)""", (combined_id, subject, inserted))
            total_inserted += inserted
            conn.commit()

        except Exception as e:
            errors.append(f"{subject}: {str(e)}")

    if errors and total_inserted == 0:
        # Complete failure — clean up
        c.execute("DELETE FROM combined_tests WHERE id=?", (combined_id,))
        conn.commit()
        conn.close()
        return jsonify({'status':'error','message':'Generation failed:\n' + '\n'.join(errors)}), 500

    conn.close()
    msg = f'Combined test "{test_no}" created with {total_inserted} questions across {len(subjects)} subjects.'
    if errors:
        msg += f'\nWarnings: ' + '; '.join(errors)
    return jsonify({'status':'success','message':msg,'combined_test_id':combined_id})

@app.route('/api/check_test_no')
def api_check_test_no():
    """Feature #5: Check if a test_no already exists for a subject (across all sections)."""
    subject = request.args.get('subject','').strip()
    test_no = request.args.get('test_no','').strip()
    class_  = request.args.get('class','').strip()
    if not subject or not test_no:
        return jsonify({'exists': False, 'details': []})
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT DISTINCT tp.class, tp.section, tp.test_no, t.name as teacher_name, tp.created_at
                 FROM test_papers tp
                 LEFT JOIN teachers t ON tp.uploaded_by = t.id AND tp.uploader_type='teacher'
                 WHERE tp.subject=? AND tp.test_no=? AND tp.class=?""",
             (subject, test_no, class_))
    rows = [dict(r) for r in c.fetchall()]
    return jsonify({'exists': len(rows) > 0, 'details': rows})


# ── MULTI-SECTION TEST CREATION ROUTE ──────────────────────────────────────────

@app.route('/api/create_multi_section_test', methods=['POST'])
@limiter.limit("15 per minute")
def api_create_multi_section_test():
    """Create a Multi-Section Test with independent per-section subjects, counts, marking schemes, and question sources."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401

    uploader_type = 'admin' if is_admin else 'teacher'
    uploader_id = str(session.get('teacher_id') or 'admin')
    teacher_name = session.get('teacher_name') or uploader_type.capitalize()

    if request.is_json:
        data = request.get_json() or {}
        board = (data.get('board') or 'CBSE').strip() or 'CBSE'
        class_ = (data.get('class') or '').strip()
        section = (data.get('section') or '').strip()
        title = (data.get('title') or '').strip()
        test_no = (data.get('test_no') or '').strip()
        timing_mode = (data.get('timing_mode') or 'disabled').strip()
        shuffle_sections = 1 if data.get('shuffle_sections') in (True, 1, '1', 'true', 'on') else 0
        remark = (data.get('remark') or '').strip()
        sections = data.get('sections', [])
    else:
        board = request.form.get('board', 'CBSE').strip() or 'CBSE'
        class_ = request.form.get('class', '').strip()
        section = request.form.get('section', '').strip()
        title = request.form.get('title', '').strip()
        test_no = request.form.get('test_no', '').strip()
        timing_mode = request.form.get('timing_mode', 'disabled').strip()
        shuffle_sections = 1 if request.form.get('shuffle_sections') in ('true', '1', 'on') else 0
        remark = request.form.get('remark', '').strip()
        try:
            sections = json.loads(request.form.get('sections_json', '[]'))
        except Exception as e:
            return jsonify({'status': 'error', 'message': f'Invalid sections data: {str(e)}'}), 400

    if not (class_ and title and test_no and sections):
        return jsonify({'status': 'error', 'message': 'Class, Title, Test No and at least one section are required'}), 400

    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()

    conn = get_db()
    c = conn.cursor()

    # Check for duplicate test title + test_no in this class
    c.execute("""SELECT id FROM test_papers 
                 WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=?""",
              (class_, f"%{clean_cls}%", title, test_no))
    if c.fetchone():
        conn.close()
        return jsonify({'status': 'error', 'message': f'A test with Title "{title}" and Test No "{test_no}" already exists in Class {class_}. Please choose a different Test No or Title.'}), 400

    # 1. Insert test_papers record
    c.execute("""INSERT INTO test_papers
                 (filename, class, section, subject, test_no, uploaded_by, uploader_type, question_count, is_active, board, paper_type)
                 VALUES (?, ?, ?, ?, ?, ?, ?, 0, 1, ?, 'multi_section')""",
              (f"{class_}{section}_{title}_{test_no}", class_, section, title, test_no, uploader_id, uploader_type, board))
    test_paper_id = c.lastrowid

    # 2. Insert test_section_config
    c.execute("""INSERT INTO test_section_config
                 (test_paper_id, timing_mode, shuffle_sections, two_stage_instructions)
                 VALUES (?, ?, ?, 1)""",
              (test_paper_id, timing_mode, shuffle_sections))

    total_inserted = 0
    errors = []

    # 3. Process each section
    for idx, sec in enumerate(sections):
        sec_name = sec.get('section_name', f'Section {chr(65 + idx)}').strip()
        sec_subj = sec.get('subject', '').strip()
        sec_count = int(sec.get('question_count', 5))
        pos_marks = float(sec.get('positive_marks', 1.0))
        neg_marks = float(sec.get('negative_marks', 0.0))
        duration_mins = int(sec.get('duration_minutes', 0)) if timing_mode == 'fixed' else 0
        gen_mode = (sec.get('generation_mode') or sec.get('source_mode') or 'ai').strip().lower()
        sec_chapter = sec.get('chapter', '').strip()

        if not sec_subj or sec_count <= 0:
            continue

        c.execute("""INSERT INTO test_sections
                     (test_paper_id, section_name, subject, question_count, positive_marks, negative_marks, duration_minutes, order_index, generation_mode)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                  (test_paper_id, sec_name, sec_subj, sec_count, pos_marks, neg_marks, duration_mins, idx, gen_mode))
        section_id = c.lastrowid

        sec_inserted = 0

        # Mode A: Question Bank Auto-Pick
        if gen_mode == 'qb':
            qb_query = "SELECT * FROM question_bank_items WHERE approved=1 AND class=? AND subject=?"
            qb_params = [class_, sec_subj]
            if sec_chapter:
                qb_query += " AND chapter LIKE ?"
                qb_params.append(f"%{sec_chapter}%")
            qb_query += " ORDER BY RANDOM() LIMIT ?"
            qb_params.append(sec_count)
            c.execute(qb_query, qb_params)
            qb_items = [dict(r) for r in c.fetchall()]

            for qi in qb_items:
                raw_opts = qi.get('raw_options', '[]')
                try:
                    opts = json.loads(raw_opts)
                except Exception:
                    opts = []
                opt_a = opts[0] if len(opts) > 0 else 'A'
                opt_b = opts[1] if len(opts) > 1 else 'B'
                opt_c = opts[2] if len(opts) > 2 else 'C'
                opt_d = opts[3] if len(opts) > 3 else 'D'
                c.execute("""INSERT INTO questions
                             (class, subject, chapter, test_no, question_type, question, question_hi,
                              option_a, option_b, option_c, option_d, correct_answer, board, section_id, image_path)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                          (class_, sec_subj, f"{sec_name}: {sec_subj}", test_no,
                           qi.get('question_type', 'MCQ'), qi.get('raw_question_text', ''), '',
                           opt_a, opt_b, opt_c, opt_d,
                           qi.get('raw_answer', ''), board, section_id, qi.get('image_path', '')))
                sec_inserted += 1

            if sec_inserted < sec_count:
                needed = sec_count - sec_inserted
                c.execute("""SELECT * FROM questions 
                             WHERE class=? AND subject=? AND (test_no != ? OR test_no IS NULL)
                             ORDER BY RANDOM() LIMIT ?""", (class_, sec_subj, test_no, needed))
                repo_items = [dict(r) for r in c.fetchall()]
                for ri in repo_items:
                    c.execute("""INSERT INTO questions
                                 (class, subject, chapter, test_no, question_type, question, question_hi,
                                  option_a, option_a_hi, option_b, option_b_hi, option_c, option_c_hi,
                                  option_d, option_d_hi, correct_answer, board, section_id, image_path)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                              (class_, sec_subj, f"{sec_name}: {sec_subj}", test_no,
                               ri.get('question_type', 'MCQ'), ri.get('question', ''), ri.get('question_hi', ''),
                               ri.get('option_a', ''), ri.get('option_a_hi', ''),
                               ri.get('option_b', ''), ri.get('option_b_hi', ''),
                               ri.get('option_c', ''), ri.get('option_c_hi', ''),
                               ri.get('option_d', ''), ri.get('option_d_hi', ''),
                               ri.get('correct_answer', ''), board, section_id, ri.get('image_path', '')))
                    sec_inserted += 1

        # Mode B: CSV Upload for section
        elif gen_mode == 'upload':
            file_key = f'csv_file_{idx}'
            csv_file = request.files.get(file_key)
            if csv_file and csv_file.filename.endswith('.csv'):
                fname = secure_filename(csv_file.filename)
                fpath = os.path.join(app.config['UPLOAD_FOLDER'], f"ms_{uuid.uuid4().hex[:8]}_{fname}")
                csv_file.save(fpath)
                try:
                    with open(fpath, 'r', encoding='utf-8-sig') as f:
                        reader = csv.DictReader(f)
                        for r_row in reader:
                            if sec_inserted >= sec_count:
                                break
                            parsed = parse_question_csv_row(r_row)
                            if not parsed or not parsed['question']:
                                continue
                            c.execute("""INSERT INTO questions
                                         (class, subject, chapter, test_no, question_type, question,
                                          option_a, option_b, option_c, option_d, correct_answer, board, section_id)
                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                      (class_, sec_subj, f"{sec_name}: {sec_subj}", test_no,
                                       parsed.get('question_type', 'MCQ'), parsed['question'],
                                       parsed['option_a'], parsed['option_b'], parsed['option_c'], parsed['option_d'],
                                       parsed['correct_answer'], board, section_id))
                            sec_inserted += 1
                finally:
                    if os.path.exists(fpath):
                        os.remove(fpath)
            else:
                errors.append(f"{sec_name} ({sec_subj}): No valid CSV file uploaded.")

        # Mode C: AI Generate (Default)
        else:
            try:
                type_counts = [('MCQ (4 options)', sec_count)]
                q_list, gen_err = generate_cbt_questions_in_batches(
                    class_, section, sec_subj, sec_chapter or f"{title} — {sec_name}", test_no,
                    remark=remark or f"Multi-Section: {sec_name}",
                    type_counts=type_counts, board=board, max_batch_size=10
                )
                if q_list and isinstance(q_list, list):
                    for q in q_list:
                        img_path = process_cbt_question_image(q, class_, section, sec_subj, sec_chapter)
                        c.execute("""INSERT INTO questions
                                     (class, subject, chapter, test_no, question_type, question, question_hi,
                                      option_a, option_a_hi, option_b, option_b_hi, option_c, option_c_hi,
                                      option_d, option_d_hi, correct_answer, image_path, board, section_id)
                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                  (class_, sec_subj, f"{sec_name}: {sec_subj}", test_no,
                                   q.get('question_type', 'MCQ'), q.get('question', ''), q.get('question_hi', ''),
                                   q.get('option_a', ''), q.get('option_a_hi', ''),
                                   q.get('option_b', ''), q.get('option_b_hi', ''),
                                   q.get('option_c', ''), q.get('option_c_hi', ''),
                                   q.get('option_d', ''), q.get('option_d_hi', ''),
                                   q.get('correct_answer', ''), img_path, board, section_id))
                        sec_inserted += 1
                else:
                    errors.append(f"{sec_name} ({sec_subj}): AI Generation failed ({gen_err})")
            except Exception as e:
                errors.append(f"{sec_name} ({sec_subj}): {str(e)}")

        total_inserted += sec_inserted

    # Update question count in test_papers
    c.execute("UPDATE test_papers SET question_count=? WHERE id=?", (total_inserted, test_paper_id))

    # Record history
    c.execute("""INSERT INTO test_generation_history
                 (teacher_id, class, section, subject, chapter, test_no, output_mode,
                  total_questions, mcq_count, remark, board, paper_type)
                 VALUES (?, ?, ?, ?, ?, ?, 'cbt', ?, ?, ?, ?, 'multi_section')""",
              (uploader_id, class_, section, title, f"Multi-Section: {len(sections)} sections", test_no,
               total_inserted, total_inserted, remark or 'Multi-Section Test', board))

    c.execute("""INSERT INTO mcq_test_history (teacher_id, teacher_name, class, section, subject, test_no, question_count)
                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
              (uploader_id if uploader_type == 'teacher' else 0, teacher_name, class_, section, title, test_no, total_inserted))

    conn.commit()
    conn.close()

    if total_inserted == 0:
        return jsonify({'status': 'error', 'message': 'Failed to create sections:\n' + '\n'.join(errors)}), 500

    msg = f'Multi-Section Test "{title}" ({test_no}) created with {total_inserted} questions across {len(sections)} sections.'
    if errors:
        msg += '\nWarnings: ' + '; '.join(errors)

    return jsonify({
        'status': 'success',
        'message': msg,
        'test_paper_id': test_paper_id,
        'title': title,
        'test_no': test_no,
        'total_questions': total_inserted
    })


@app.route('/waiting')
@student_required
def waiting():
    """
    Render the waiting page for students to wait for exam to start.
    """
    student_id = session.get('student_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT is_active FROM exam_control WHERE id=1")
    row = c.fetchone()
    c.execute("SELECT status, class, section, subject, test_no FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()

    active_sess = get_active_session_for_student(student_id=student_id, conn=conn)
    is_active = (active_sess is not None and active_sess.get('status') in ('running', 'paused')) or (row and bool(row['is_active']))

    lock_info = None
    if student and student['class'] and student['subject']:
        tno = student['test_no'] or session.get('test_no', '')
        # Check if test paper is inactive (only if exam session is NOT actively running)
        if tno and not is_active:
            clean_cls = student['class'].replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            c.execute("""SELECT id FROM test_papers 
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? AND is_active=0""",
                      (student['class'], f"%{clean_cls}%", student['subject'], tno))
            if c.fetchone():
                lock_info = {
                    'title': 'Test Inactive',
                    'message': f"Test '{tno}' is currently marked inactive by the faculty/admin and cannot be attempted at this time.",
                    'icon': 'fa-pause-circle',
                    'color': '#d97706',
                    'state': 'inactive'
                }

        if not lock_info:
            sch_res = get_student_schedule_status(student['class'], student['section'], student['subject'], tno, student_id)
            if sch_res['state'] == 'upcoming':
                lock_info = {'title': 'Exam Locked', 'message': sch_res['message'], 'icon': 'fa-lock', 'color': '#eab308', 'state': 'upcoming'}
            elif sch_res['state'] == 'expired' and student['status'] == 'Not Started':
                lock_info = {'title': 'Exam Window Closed', 'message': sch_res['message'], 'icon': 'fa-times-circle', 'color': '#ef4444', 'state': 'expired'}
            elif sch_res['state'] == 'none':
                ms_info = _get_multi_section_info(student['class'], student['subject'], tno, conn)
                if not ms_info:
                    lock_info = {'title': 'Test Not Scheduled', 'message': f"Test '{tno}' is not scheduled. Students can only take scheduled tests.", 'icon': 'fa-calendar-times', 'color': '#ef4444', 'state': 'none'}

    active_sess = get_active_session_for_student(student_id=student_id, conn=conn)
    is_active = (active_sess is not None and active_sess.get('status') in ('running', 'paused')) or (row and bool(row['is_active']))
    conn.close()

    if not lock_info and is_active:
        if session.get('resume_mode'):
            return redirect(url_for('exam'))
        if session.get('is_multi_section'):
            return redirect(url_for('instructions_stage1'))
        return redirect(url_for('guidelines'))
    return render_template('waiting.html', lock_info=lock_info)

@app.route('/check_exam_status')
@student_required
def check_exam_status():
    """
    API endpoint to check if the exam has started.
    Returns session-specific state, paused status, and custom end_message if ended.
    """
    auto_submit_expired_exams()
    update_scheduled_tests_status()
    student_id = session.get('student_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT is_active, start_time, duration FROM exam_control WHERE id=1")
    row = c.fetchone()
    c.execute("SELECT status, class, section, subject, test_no FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()

    force_submitted = bool(student and student['status'] == 'Submitted')
    if force_submitted:
        # Check if the session ended with a custom message
        latest_sess = get_latest_session_for_student(student_id=student_id, conn=conn)
        end_message = latest_sess.get('end_message') if latest_sess else None
        conn.close()
        return jsonify({'active': False, 'force_submitted': True, 'end_message': end_message})

    if student and student['class'] and student['subject']:
        tno = student['test_no'] or session.get('test_no', '')
        if tno:
            clean_cls = student['class'].replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            c.execute("""SELECT id FROM test_papers 
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? AND is_active=0""",
                      (student['class'], f"%{clean_cls}%", student['subject'], tno))
            if c.fetchone():
                conn.close()
                return jsonify({'active': False, 'force_submitted': False, 'test_inactive': True, 'message': f"Test '{tno}' is currently inactive."})

        sch_res = get_student_schedule_status(student['class'], student['section'], student['subject'], tno, student_id)
        if sch_res['state'] == 'upcoming':
            conn.close()
            return jsonify({'active': False, 'force_submitted': False, 'schedule_locked': True, 'state': 'upcoming', 'message': sch_res['message']})
        elif sch_res['state'] == 'expired' and student['status'] == 'Not Started':
            conn.close()
            return jsonify({'active': False, 'force_submitted': False, 'schedule_locked': True, 'state': 'expired', 'message': sch_res['message']})
        elif sch_res['state'] == 'none':
            conn.close()
            return jsonify({'active': False, 'force_submitted': False, 'schedule_locked': True, 'state': 'none', 'message': f"Test '{tno}' is not scheduled."})

    # Check session-based control first
    active_sess = get_active_session_for_student(student_id=student_id, conn=conn)
    if active_sess:
        conn.close()
        is_paused = (active_sess.get('status') == 'paused')
        return jsonify({'active': True, 'paused': is_paused, 'force_submitted': False, 'session_id': active_sess['id']})

    # Fallback to global exam_control
    is_active = bool(row and row['is_active'])
    conn.close()
    return jsonify({'active': is_active, 'paused': False, 'force_submitted': False})

@app.route('/exam')
@student_required
def exam():
    """
    Render the exam page with shuffled questions for the student.
    """
    if session.get('test_type') == 'Coding Test':
        return redirect(url_for('coding_exam_list'))

    import random
    conn = get_db()
    c = conn.cursor()

    active_sess = get_active_session_for_student(student_id=session.get('student_id'), conn=conn)
    is_active = (active_sess is not None and active_sess.get('status') in ('running', 'paused'))
    if not is_active:
        c.execute("SELECT is_active FROM exam_control WHERE id=1")
        row = c.fetchone()
        if not row or not row['is_active']:
            conn.close()
            return redirect(url_for('waiting'))

    student_id = session['student_id']
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if student and student['status'] == 'Submitted':
        conn.close()
        return redirect(url_for('waiting'))

    if student and student['status'] == 'Not Started':
        tno = (student['test_no'] if ('test_no' in student.keys() and student['test_no']) else session.get('test_no', '') or '').strip()
        if tno:
            clean_cls = student['class'].replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            c.execute("""SELECT id FROM test_papers 
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? AND is_active=0""",
                      (student['class'], f"%{clean_cls}%", student['subject'], tno))
            if c.fetchone():
                conn.close()
                return redirect(url_for('waiting'))

        ms_info = _get_multi_section_info(student['class'], student['subject'], tno, conn)
        sch_res = get_student_schedule_status(student['class'], student['section'], student['subject'], tno, student_id)
        if sch_res['state'] in ('upcoming', 'expired'):
            conn.close()
            return redirect(url_for('waiting'))
        elif sch_res['state'] == 'none' and not ms_info:
            conn.close()
            return redirect(url_for('waiting'))

    if student:
        tno = (student['test_no'] if ('test_no' in student.keys() and student['test_no']) else session.get('test_no', '') or '').strip()
        ms_info = _get_multi_section_info(student['class'], student['subject'], tno, conn)

        if ms_info:
            c.execute("""SELECT COUNT(q.id) FROM questions q
                         JOIN test_sections ts ON q.section_id = ts.id
                         WHERE ts.test_paper_id = ?""", (ms_info['tp_id'],))
            q_count = c.fetchone()[0]
        else:
            c.execute("SELECT COUNT(*) FROM questions WHERE class=? AND subject=?",
                      (student['class'], student['subject']))
            q_count = c.fetchone()[0]

        if q_count == 0:
            conn.close()
            return "No questions available for this subject/class. Please contact admin.", 404

        c.execute("SELECT COUNT(*) FROM shuffled_questions WHERE student_id=?", (student_id,))
        count = c.fetchone()[0]
        if count == 0:
            test_no = (student['test_no'] if ('test_no' in student.keys() and student['test_no']) else session.get('test_no', '') or '').strip()
            if ms_info:
                # Multi-section: populate across all sections
                c.execute("""SELECT id, section_name, subject, duration_minutes, order_index
                             FROM test_sections WHERE test_paper_id=? ORDER BY order_index ASC""", (ms_info['tp_id'],))
                sec_rows = [dict(r) for r in c.fetchall()]
                global_q_idx = 0
                first_sec_id = None
                first_sec_name = ''
                first_sec_dur = 0
                for s_idx, sec in enumerate(sec_rows):
                    sec_display_name = sec['section_name']
                    if sec.get('subject') and sec.get('subject').lower() not in sec_display_name.lower():
                        sec_display_name = f"{sec['section_name']} - {sec['subject']}"
                    if s_idx == 0:
                        first_sec_id = sec['id']
                        first_sec_name = sec_display_name
                        first_sec_dur = sec['duration_minutes'] or 0
                    c.execute("SELECT id FROM questions WHERE section_id=? ORDER BY id", (sec['id'],))
                    s_qids = [r['id'] for r in c.fetchall()]
                    if not s_qids:
                        c.execute("SELECT id FROM questions WHERE class=? AND subject=? AND test_no=? ORDER BY id",
                                  (student['class'], sec['subject'], test_no))
                        s_qids = [r['id'] for r in c.fetchall()]
                    random.shuffle(s_qids)
                    for qid in s_qids:
                        opt_order = _shuffle_options_safe(qid, c)
                        c.execute("INSERT INTO shuffled_questions (student_id, question_id, shuffled_index, option_order) VALUES (?,?,?,?)",
                                  (student_id, qid, global_q_idx, opt_order))
                        global_q_idx += 1
                c.execute("""INSERT OR REPLACE INTO student_section_progress
                             (id, student_id, test_paper_id, current_section_id, locked_sections, section_time_left, updated_at)
                             VALUES (
                                 (SELECT id FROM student_section_progress WHERE student_id=? AND test_paper_id=?),
                                 ?, ?, ?, '[]', ?, CURRENT_TIMESTAMP
                             )""",
                          (student_id, ms_info['tp_id'], student_id, ms_info['tp_id'], first_sec_id, first_sec_dur * 60))
                if first_sec_name:
                    c.execute("UPDATE students SET current_section=? WHERE student_id=?", (first_sec_name, student_id))
            elif test_no:
                c.execute("SELECT id FROM questions WHERE class=? AND subject=? AND (test_no=? OR chapter=?) ORDER BY id",
                          (student['class'], student['subject'], test_no, test_no))
                qids = [row['id'] for row in c.fetchall()]
                random.shuffle(qids)
                for idx, qid in enumerate(qids):
                    opt_order = _shuffle_options_safe(qid, c)
                    c.execute("INSERT INTO shuffled_questions (student_id, question_id, shuffled_index, option_order) VALUES (?,?,?,?)",
                              (student_id, qid, idx, opt_order))
            else:
                c.execute("SELECT id FROM questions WHERE class=? AND subject=? ORDER BY id",
                          (student['class'], student['subject']))
                qids = [row['id'] for row in c.fetchall()]
                random.shuffle(qids)
                for idx, qid in enumerate(qids):
                    opt_order = _shuffle_options_safe(qid, c)
                    c.execute("INSERT INTO shuffled_questions (student_id, question_id, shuffled_index, option_order) VALUES (?,?,?,?)",
                              (student_id, qid, idx, opt_order))
            conn.commit()

        if student['exam_started_at'] is None or student['status'] == 'Not Started':
            now_str = datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
            c.execute("UPDATE students SET status='In Progress', exam_started_at=? WHERE student_id=?",
                      (now_str, student_id))
            conn.commit()

    # Get chapter/test_no for header
    c.execute("SELECT DISTINCT chapter FROM questions WHERE class=? AND subject=? AND chapter IS NOT NULL AND chapter!='' LIMIT 1",
              (student['class'], student['subject']))
    chapter_row = c.fetchone()
    test_no = chapter_row['chapter'] if chapter_row else ''

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path   = get_setting('logo_path', '')
    conn.close()

    now = datetime.datetime.now()
    combined_subjects = session.get('combined_subjects', [])
    is_combined       = bool(session.get('combined_test_id'))
    is_multi_section  = bool(session.get('is_multi_section') or ms_info)
    timing_mode       = ms_info['timing_mode'] if ms_info else session.get('timing_mode', 'disabled')
    test_paper_id     = ms_info['tp_id'] if ms_info else session.get('test_paper_id')

    return render_template('exam.html',
                           school_name=school_name,
                           logo_path=logo_path,
                           gemini_key=_get_gemini_key(),
                           student_name=session.get('student_name',''),
                           student_id=student_id,
                           student_class=student['class'],
                           student_section=student['section'] if student['section'] else session.get('section', ''),
                           student_subject=student['subject'],
                           test_no=test_no,
                           exam_date=now.strftime('%d/%m/%Y'),
                           exam_day=now.strftime('%A'),
                           is_combined=is_combined,
                           combined_subjects=combined_subjects,
                           is_multi_section=is_multi_section,
                           timing_mode=timing_mode,
                           test_paper_id=test_paper_id)

# ── Language Subject Detection Helper ───────────────────────────────────────────
def is_language_subject(subj):
    """Return True if the subject represents a language subject (English, Hindi, Sanskrit, etc.)."""
    if not subj:
        return False
    s = str(subj).strip().lower()
    lang_keywords = [
        'english', 'अंग्रेजी', 'अंग्रेज़ी',
        'hindi', 'हिन्दी', 'हिंदी',
        'sanskrit', 'संस्कृत',
        'urdu', 'उर्दू',
        'french', 'français',
        'german', 'deutsch',
        'spanish', 'español',
        'bengali', 'bangla', 'বাংলা',
        'marathi', 'मराठी',
        'gujarati', 'ગુજરાતી',
        'tamil', 'தமிழ்',
        'telugu', 'తెలుగు',
        'kannada', 'ಕನ್ನಡ',
        'malayalam', 'മലയാളം',
        'punjabi', 'ਪੰਜਾਬੀ',
        'odia', 'oriya', 'ଓଡ଼ିଆ',
        'assamese', 'অসমীয়া',
        'maithili', 'मैथिली',
        'kashmiri', 'कॉशुर',
        'nepali', 'नेपाली',
        'konkani', 'कोंकणी',
        'sindhi', 'سنڌي',
        'manipuri', 'meitei',
        'arabic', 'عربي',
        'persian', 'فارسی',
        'russian', 'русский',
        'japanese', '日本語',
        'chinese', 'mandarin', '中文',
        'korean', '한국어',
        'latin',
        'language', 'literature', 'vyakaran', 'sahitya', 'grammar', 'bhasha'
    ]
    return any(kw in s for kw in lang_keywords)

@app.route('/get_questions')
@student_required
def get_questions():
    """
    API endpoint to retrieve exam questions with shuffled options for display.
    """
    student_id = session.get('student_id')
    class_     = session.get('class')
    subject    = session.get('subject')
    conn = get_db()
    c    = conn.cursor()

    c.execute('''
        SELECT q.id, q.question, q.option_a, q.option_b, q.option_c, q.option_d,
               q.question_hi, q.option_a_hi, q.option_b_hi, q.option_c_hi, q.option_d_hi,
               q.image_path, q.subject, q.question_type, q.section_id,
               COALESCE(sq.option_order, 'ABCD') as option_order,
               r.selected_option,
               ts.section_name, ts.positive_marks, ts.negative_marks, ts.duration_minutes, ts.order_index
        FROM questions q
        LEFT JOIN shuffled_questions sq ON q.id = sq.question_id AND sq.student_id = ?
        LEFT JOIN responses r           ON q.id = r.question_id  AND r.student_id = ?
        LEFT JOIN test_sections ts      ON q.section_id = ts.id
        WHERE sq.student_id = ?
        ORDER BY COALESCE(sq.shuffled_index, q.id)
    ''', (student_id, student_id, student_id))
    rows = c.fetchall()
    conn.close()

    if not rows:
        return jsonify([])

    questions = []
    for row in rows:
        opt_order     = row['option_order']
        original_opts = {'A': row['option_a'], 'B': row['option_b'],
                         'C': row['option_c'], 'D': row['option_d']}
        original_opts_hi = {'A': row['option_a_hi'] or '', 'B': row['option_b_hi'] or '',
                            'C': row['option_c_hi'] or '', 'D': row['option_d_hi'] or ''}
        shuffled_options = {}
        shuffled_options_hi = {}
        for display_idx, orig_letter in enumerate(opt_order):
            display_letter = chr(ord('A') + display_idx)
            shuffled_options[display_letter] = original_opts[orig_letter]
            shuffled_options_hi[display_letter] = original_opts_hi[orig_letter]

        raw_image = row['image_path'] or ''
        if raw_image.startswith('__smiles__'):
            smiles_val   = raw_image[len('__smiles__'):]
            actual_image = None
        else:
            smiles_val   = ''
            actual_image = raw_image if raw_image else None

        q_subj = row['subject'] or ''
        is_lang = is_language_subject(q_subj)
        is_pure_hindi = any(h in q_subj.lower() for h in ['hindi', 'हिन्दी', 'हिंदी'])

        # Language subjects other than pure Hindi MUST NEVER have Hindi translation fields
        if is_lang and not is_pure_hindi:
            q_hi = ''
            opt_a_hi = ''
            opt_b_hi = ''
            opt_c_hi = ''
            opt_d_hi = ''
        else:
            q_hi = row['question_hi'] or ''
            opt_a_hi = shuffled_options_hi['A']
            opt_b_hi = shuffled_options_hi['B']
            opt_c_hi = shuffled_options_hi['C']
            opt_d_hi = shuffled_options_hi['D']

        questions.append({
            'id':               row['id'],
            'question':         row['question'],
            'question_hi':      q_hi,
            'option_a':         shuffled_options['A'],
            'option_b':         shuffled_options['B'],
            'option_c':         shuffled_options['C'],
            'option_d':         shuffled_options['D'],
            'option_a_hi':      opt_a_hi,
            'option_b_hi':      opt_b_hi,
            'option_c_hi':      opt_c_hi,
            'option_d_hi':      opt_d_hi,
            'image_path':       actual_image,
            'smiles':           smiles_val,
            'subject':          row['subject'],       # for tab grouping
            'question_type':    row['question_type'] or 'MCQ',
            'selected':         row['selected_option'] or '',
            'section_id':       row['section_id'] or None,
            'section_name':     row['section_name'] or '',
            'positive_marks':   float(row['positive_marks']) if row['positive_marks'] is not None else 1.0,
            'negative_marks':   float(row['negative_marks']) if row['negative_marks'] is not None else 0.0,
            'duration_minutes': int(row['duration_minutes']) if row['duration_minutes'] is not None else 0,
            'order_index':      row['order_index'] if row['order_index'] is not None else 0,
        })
    return jsonify(questions)

@app.route('/api/section_progress', methods=['GET', 'POST'])
@student_required
def api_section_progress():
    """Retrieve or update student section progress, locks, and timing."""
    student_id = session.get('student_id')
    test_paper_id = session.get('test_paper_id')
    timing_mode = session.get('timing_mode', 'disabled')

    conn = get_db()
    c = conn.cursor()

    if not test_paper_id:
        c.execute("SELECT class, subject, test_no FROM students WHERE student_id=?", (student_id,))
        st = c.fetchone()
        if st:
            ms_info = _get_multi_section_info(st['class'], st['subject'], st['test_no'], conn)
            if ms_info:
                test_paper_id = ms_info['tp_id']
                timing_mode = ms_info['timing_mode']
                session['test_paper_id'] = test_paper_id
                session['timing_mode'] = timing_mode

    if not test_paper_id:
        conn.close()
        return jsonify({'status': 'none', 'message': 'Not a multi-section test'}), 200

    if request.method == 'POST':
        data = request.get_json() or {}
        curr_sec_id = data.get('current_section_id')
        lock_sec_id = data.get('lock_section_id')
        sec_time_left = data.get('section_time_left')

        c.execute("SELECT * FROM student_section_progress WHERE student_id=? AND test_paper_id=?",
                  (student_id, test_paper_id))
        prog = c.fetchone()
        locked = []
        if prog and prog['locked_sections']:
            try:
                locked = json.loads(prog['locked_sections'])
            except Exception:
                locked = []

        if lock_sec_id and lock_sec_id not in locked:
            locked.append(lock_sec_id)

        locked_json = json.dumps(locked)
        time_left_val = int(sec_time_left) if sec_time_left is not None else (prog['section_time_left'] if prog else 0)
        new_curr_id = curr_sec_id if curr_sec_id is not None else (prog['current_section_id'] if prog else None)

        c.execute("""INSERT OR REPLACE INTO student_section_progress
                     (id, student_id, test_paper_id, current_section_id, locked_sections, section_time_left, updated_at)
                     VALUES (
                         (SELECT id FROM student_section_progress WHERE student_id=? AND test_paper_id=?),
                         ?, ?, ?, ?, ?, CURRENT_TIMESTAMP
                     )""",
                  (student_id, test_paper_id, student_id, test_paper_id, new_curr_id, locked_json, time_left_val))

        if new_curr_id:
            c.execute("SELECT section_name, subject FROM test_sections WHERE id=?", (new_curr_id,))
            sec_name_row = c.fetchone()
            if sec_name_row:
                s_name = sec_name_row['section_name']
                s_subj = sec_name_row['subject']
                if s_subj and s_subj.lower() not in s_name.lower():
                    s_name = f"{s_name} - {s_subj}"
                c.execute("UPDATE students SET current_section=? WHERE student_id=?", (s_name, student_id))

        conn.commit()
        conn.close()
        return jsonify({'status': 'ok', 'locked_sections': locked, 'current_section_id': new_curr_id})

    # GET
    c.execute("SELECT * FROM student_section_progress WHERE student_id=? AND test_paper_id=?",
              (student_id, test_paper_id))
    prog = c.fetchone()
    locked = []
    curr_sec_id = None
    time_left = 0
    if prog:
        curr_sec_id = prog['current_section_id']
        time_left = prog['section_time_left']
        try:
            locked = json.loads(prog['locked_sections'] or '[]')
        except Exception:
            locked = []

    c.execute("""SELECT id, section_name, subject, question_count, positive_marks, negative_marks, duration_minutes, order_index
                 FROM test_sections WHERE test_paper_id=? ORDER BY order_index ASC""", (test_paper_id,))
    secs = [dict(r) for r in c.fetchall()]

    conn.close()
    return jsonify({
        'status': 'ok',
        'is_multi_section': True,
        'timing_mode': timing_mode,
        'current_section_id': curr_sec_id,
        'locked_sections': locked,
        'section_time_left': time_left,
        'sections': secs
    })

@app.route('/save_answer', methods=['POST'])
@student_required
@validate_schema(SAVE_ANSWER_SCHEMA, is_json=True)
def save_answer():
    """
    API endpoint to save a student's answer to a question.
    """
    data = request.get_json()
    question_id = data.get('question_id')
    displayed_option = data.get('selected_option')
    student_id = session['student_id']

    conn = get_db()
    c = conn.cursor()
    # Get the option order for this question
    c.execute("SELECT option_order FROM shuffled_questions WHERE student_id=? AND question_id=?",
              (student_id, question_id))
    row = c.fetchone()
    original_letter = ''
    if row and displayed_option:
        opt_order = row['option_order']
        display_index = ord(displayed_option) - ord('A')
        if display_index < len(opt_order):
            original_letter = opt_order[display_index]

    c.execute("INSERT OR REPLACE INTO responses (student_id, question_id, selected_option, created_at) VALUES (?,?,?, CURRENT_TIMESTAMP)",
              (student_id, question_id, original_letter))
    c.execute("UPDATE students SET status='In Progress' WHERE student_id=?", (student_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'ok'})

@app.route('/submit_exam', methods=['POST'])
@student_required
def submit_exam():
    """
    API endpoint to submit the exam and calculate results.
    Accepts optional submission reason (e.g. 'tab_switch', 'fullscreen_exit').
    """
    student_id        = session['student_id']
    class_            = session.get('class')
    subject           = session.get('subject')
    name              = session.get('student_name')
    combined_subjects = session.get('combined_subjects', [])
    test_no           = session.get('test_no', '')

    reason = 'manual'
    if request.is_json:
        reason = request.json.get('reason', 'manual')
    elif request.form.get('reason'):
        reason = request.form.get('reason', 'manual')

    session['submission_reason'] = reason

    conn = get_db()
    c    = conn.cursor()

    c.execute("SELECT negative_marking, negative_value FROM exam_control WHERE id=1")
    ec          = c.fetchone()
    neg_enabled = bool(ec['negative_marking']) if ec else False
    neg_value   = float(ec['negative_value'])  if ec else 0.33

    def _score_subject(subj):
        c.execute("""SELECT q.correct_answer, q.option_a, q.option_b, q.option_c, q.option_d, r.selected_option, q.subject
                     FROM shuffled_questions sq
                     JOIN questions q   ON q.id = sq.question_id
                     LEFT JOIN responses r ON q.id = r.question_id AND r.student_id = ?
                     WHERE sq.student_id = ? AND q.subject = ?""",
                  (student_id, student_id, subj))
        rows  = c.fetchall()
        total = len(rows)
        raw   = 0.0
        for r in rows:
            sel = (r['selected_option'] or '').strip().upper()
            cor_letter = _normalize_correct_answer(
                r['correct_answer'],
                r['option_a'],
                r['option_b'],
                r['option_c'],
                r['option_d']
            )
            if sel and cor_letter and sel == cor_letter:
                raw += 1.0
            elif sel and neg_enabled:
                raw -= neg_value
        score = max(0.0, raw)
        pct   = round((score / total * 100), 2) if total > 0 else 0.0
        return score, total, pct

    c.execute("SELECT section, test_no FROM students WHERE student_id=?", (student_id,))
    st_row = c.fetchone()
    section_val = (st_row['section'] if st_row and st_row['section'] else session.get('section', ''))
    test_no_val = (st_row['test_no'] if st_row and st_row['test_no'] else session.get('test_no', ''))

    ms_info = _get_multi_section_info(class_, subject, test_no_val, conn)
    if ms_info:
        tp_id = ms_info['tp_id']
        c.execute("""SELECT id, section_name, subject, question_count, positive_marks, negative_marks, duration_minutes, order_index
                     FROM test_sections WHERE test_paper_id=? ORDER BY order_index ASC""", (tp_id,))
        sections = [dict(r) for r in c.fetchall()]

        total_exam_questions = 0
        total_exam_score = 0.0
        max_possible_marks = 0.0
        section_breakdowns = []

        for sec in sections:
            sec_id = sec['id']
            pos_m = float(sec['positive_marks'] if sec['positive_marks'] is not None else 1.0)
            neg_m = float(sec['negative_marks'] if sec['negative_marks'] is not None else 0.0)

            c.execute("""SELECT q.id, q.correct_answer, q.option_a, q.option_b, q.option_c, q.option_d, r.selected_option
                         FROM shuffled_questions sq
                         JOIN questions q ON q.id = sq.question_id
                         LEFT JOIN responses r ON q.id = r.question_id AND r.student_id = ?
                         WHERE sq.student_id = ? AND q.section_id = ?""",
                      (student_id, student_id, sec_id))
            q_rows = c.fetchall()
            sec_total = len(q_rows)
            sec_attempted = 0
            sec_correct = 0
            sec_incorrect = 0
            sec_score = 0.0
            sec_max = sec_total * pos_m

            for qr in q_rows:
                sel = (qr['selected_option'] or '').strip().upper()
                cor_letter = _normalize_correct_answer(
                    qr['correct_answer'], qr['option_a'], qr['option_b'], qr['option_c'], qr['option_d']
                )
                if sel:
                    sec_attempted += 1
                    if cor_letter and sel == cor_letter:
                        sec_correct += 1
                        sec_score += pos_m
                    else:
                        sec_incorrect += 1
                        sec_score -= neg_m

            sec_score_final = max(0.0, sec_score)
            total_exam_questions += sec_total
            total_exam_score += sec_score_final
            max_possible_marks += sec_max

            section_breakdowns.append({
                'section_id': sec_id,
                'section_name': sec['section_name'],
                'subject': sec['subject'],
                'total': sec_total,
                'attempted': sec_attempted,
                'correct': sec_correct,
                'incorrect': sec_incorrect,
                'marks_obtained': round(sec_score_final, 2),
                'max_marks': round(sec_max, 2),
                'positive_marks': pos_m,
                'negative_marks': neg_m
            })

        overall_pct = round((total_exam_score / max_possible_marks * 100), 2) if max_possible_marks > 0 else (
            round((total_exam_score / total_exam_questions * 100), 2) if total_exam_questions > 0 else 0.0
        )
        breakdown_json = json.dumps(section_breakdowns)

        _save_or_update_result(c, student_id, name, class_, section_val, subject,
                               round(total_exam_score, 2), total_exam_questions, overall_pct, test_no_val,
                               section_breakdown=breakdown_json)
    elif combined_subjects:
        for subj in combined_subjects:
            score, total, pct = _score_subject(subj)
            _save_or_update_result(c, student_id, name, class_, section_val, subj, score, total, pct, test_no_val)
    else:
        score, total, pct = _score_subject(subject)
        _save_or_update_result(c, student_id, name, class_, section_val, subject, score, total, pct, test_no_val)

    c.execute("UPDATE students SET status='Submitted' WHERE student_id=?", (student_id,))
    conn.commit()
    conn.close()

    if reason in ('tab_switch', 'fullscreen_exit'):
        log_audit_event('student', student_id, f'EXAM_AUTO_SUBMIT_{reason.upper()}', 'students', student_id, request.remote_addr)

    return jsonify({'status': 'submitted', 'redirect': url_for('submitted')})

@app.route('/submitted')
@student_required
def submitted():
    """
    Render the page shown after exam submission with reattempt option and security reason display.
    """
    student_id = session.get('student_id')
    class_ = session.get('class')
    subject = session.get('subject')
    reason = session.get('submission_reason', 'manual')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT status FROM reattempt_requests WHERE student_id=? AND class=? AND subject=? AND status='pending'",
              (student_id, class_, subject))
    pending = c.fetchone()
    conn.close()
    return render_template('submitted.html', pending=bool(pending), reason=reason)

@app.route('/request_reattempt', methods=['POST'])
@student_required
@validate_schema(REATTEMPT_REQUEST_SCHEMA, is_json=True)
def request_reattempt():
    """
    API endpoint for students to request reattempt of an exam.
    """
    student_id = session.get('student_id')
    class_ = session.get('class')
    subject = session.get('subject')

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM reattempt_requests WHERE student_id=? AND class=? AND subject=? AND status='pending'",
              (student_id, class_, subject))
    if c.fetchone():
        conn.close()
        return jsonify({'status': 'error', 'message': 'You already have a pending request.'}), 400

    c.execute("INSERT INTO reattempt_requests (student_id, class, subject, status) VALUES (?,?,?, 'pending')",
              (student_id, class_, subject))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Request sent to admin.'})

@app.route('/check_reattempt_status')
@student_required
def check_reattempt_status():
    """
    API endpoint to check the status of a reattempt request.
    """
    student_id = session.get('student_id')
    class_ = session.get('class')
    subject = session.get('subject')

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM reattempt_requests WHERE student_id=? AND class=? AND subject=? AND status='approved'",
              (student_id, class_, subject))
    approved = c.fetchone()
    c.execute("SELECT is_active FROM exam_control WHERE id=1")
    exam_row = c.fetchone()
    exam_active = bool(exam_row['is_active']) if exam_row else False
    c.execute("SELECT status FROM students WHERE student_id=?", (student_id,))
    student_row = c.fetchone()
    student_status = student_row['status'] if student_row else 'Unknown'
    conn.close()

    can_start = (approved is not None) and exam_active and (student_status == 'Not Started')
    return jsonify({
        'approved': approved is not None,
        'exam_active': exam_active,
        'student_status': student_status,
        'can_start': can_start
    })

@app.route('/get_exam_time')
@student_required
def get_exam_time():
    """
    Session-aware timer — calculates countdown based on the student's active exam session.
    Accounts for paused status and total paused time.
    Falls back to scheduled test or global exam_control if no session is active.
    """
    student_id = session.get('student_id')
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT class, section, subject, test_no, exam_started_at, status FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()

    # 1. Check if student belongs to an active exam session
    sess = get_active_session_for_student(student_id=student_id, conn=conn)

    if sess:
        duration = int(sess.get('allotted_duration') or 60)
        is_paused = (sess.get('status') == 'paused')
        sess_remaining = get_session_remaining_seconds(sess)

        if student and student['status'] == 'Submitted':
            conn.close()
            return jsonify({'remaining': 0, 'duration': duration, 'paused': False, 'session_id': sess['id']})

        if is_paused:
            conn.close()
            return jsonify({
                'remaining': sess_remaining,
                'duration': duration,
                'paused': True,
                'session_id': sess['id'],
                'message': 'Exam is currently paused by proctor.'
            })

        # Running session: student countdown
        if not student or not student['exam_started_at']:
            remaining = sess_remaining
        else:
            try:
                started_dt = parse_session_dt(student['exam_started_at'])
                now = datetime.datetime.utcnow()
                tot_paused = int(sess.get('total_paused_secs') or 0)
                elapsed = (now - started_dt).total_seconds() - tot_paused
                std_rem = int((duration * 60) - max(0, elapsed))
                remaining = max(0, min(std_rem, sess_remaining))
            except Exception:
                remaining = sess_remaining

        conn.close()
        return jsonify({'remaining': remaining, 'duration': duration, 'paused': False, 'session_id': sess['id']})

    # 2. Fallback to scheduled tests & global exam_control
    c.execute("SELECT duration, is_active FROM exam_control WHERE id=1")
    ctrl = c.fetchone()
    if not ctrl or not ctrl['is_active']:
        conn.close()
        return jsonify({'remaining': 0, 'duration': 0, 'paused': False})

    duration = ctrl['duration']
    if student and student['class'] and student['subject']:
        tno = student['test_no'] or session.get('test_no', '')
        sch_res = get_student_schedule_status(student['class'], student['section'], student['subject'], tno, student_id)
        if sch_res['state'] == 'active' and 'effective_duration' in sch_res:
            duration = sch_res['effective_duration']
        elif sch_res.get('schedule'):
            duration = sch_res['schedule']['duration_minutes']

    if not student or not student['exam_started_at']:
        conn.close()
        return jsonify({'remaining': duration * 60, 'duration': duration, 'paused': False})

    if student['status'] == 'Submitted':
        conn.close()
        return jsonify({'remaining': 0, 'duration': duration, 'paused': False})

    try:
        started_str = student['exam_started_at'].replace('T', ' ')
        start = datetime.datetime.fromisoformat(started_str)
    except Exception:
        conn.close()
        return jsonify({'remaining': duration * 60, 'duration': duration, 'paused': False})

    end_time  = start + datetime.timedelta(minutes=duration)
    remaining = int((end_time - datetime.datetime.utcnow()).total_seconds())
    if remaining < 0:
        remaining = 0

    conn.close()
    return jsonify({'remaining': remaining, 'duration': duration, 'paused': False})

# -------------------------------
# Admin Routes
# -------------------------------
@app.route('/admin', methods=['GET', 'POST'])
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    """
    Handle admin login and render admin login page.
    Supports /admin and /admin/login routes with environment-configured credentials.
    """
    if session.get('admin_logged_in'):
        return redirect(url_for('admin_dashboard'))

    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        
        admin_user = os.environ.get('ADMIN_USERNAME', 'admin').strip()
        admin_pass = os.environ.get('ADMIN_PASSWORD', 'admin123').strip()

        # If username is omitted in the login form, default to configured admin_user
        if not username:
            username = admin_user

        if (username.lower() == admin_user.lower() or username.lower() == 'admin') and password == admin_pass:
            session['admin_logged_in'] = True
            session['admin_username'] = username
            log_audit_event('admin', 'admin', 'LOGIN_SUCCESS', 'admin', 'admin', request.remote_addr)
            return redirect(url_for('admin_dashboard'))
        else:
            log_audit_event('admin', username or 'unknown', 'LOGIN_FAILED', 'admin', 'admin', request.remote_addr)
            return render_template('admin_login.html', error='Invalid credentials')
    return render_template('admin_login.html')

@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    """
    Render the admin dashboard with statistics.
    """
    conn = get_db()
    c = conn.cursor()

    # ── Existing stats ──────────────────────────────────────────────
    c.execute("SELECT COUNT(*) FROM students WHERE status='In Progress'")
    active_students = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM teachers WHERE status='active'")
    active_teachers = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM reattempt_requests WHERE status='pending'")
    pending_reattempts = c.fetchone()[0]

    c.execute("SELECT COUNT(*) FROM students")
    total_students = c.fetchone()[0]

    # ── NEW: Class, Subject, Test counts ──────────────────────────
    c.execute("SELECT COUNT(DISTINCT class) FROM questions WHERE class IS NOT NULL AND class != ''")
    total_classes = c.fetchone()[0] or 0

    c.execute("SELECT COUNT(DISTINCT subject) FROM questions WHERE subject IS NOT NULL AND subject != ''")
    total_subjects = c.fetchone()[0] or 0

    c.execute("SELECT COUNT(*) FROM test_papers")
    total_tests = c.fetchone()[0] or 0

    # ── Breakdowns ──────────────────────────────────────────────────
    c.execute("""
        SELECT class, COUNT(*) 
        FROM students 
        WHERE class IS NOT NULL AND class != '' 
        GROUP BY class 
        ORDER BY class
    """)
    class_student_counts = {row[0]: row[1] for row in c.fetchall()}

    c.execute("""
        SELECT subject, COUNT(*) 
        FROM test_papers 
        GROUP BY subject 
        ORDER BY subject
    """)
    subject_test_counts = {row[0]: row[1] for row in c.fetchall()}

    # ── Class → Subject map (for dropdowns) ──────────────────────
    c.execute("SELECT DISTINCT class, subject FROM questions WHERE class IS NOT NULL AND class != '' ORDER BY class, subject")
    rows = c.fetchall()
    class_subject_map = {}
    for row in rows:
        class_subject_map.setdefault(row['class'], []).append(row['subject'])

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')

    conn.close()

    return render_template('admin_dashboard.html',
                           active_students=active_students,
                           active_teachers=active_teachers,
                           pending_reattempts=pending_reattempts,
                           total_students=total_students,
                           total_classes=total_classes,
                           total_subjects=total_subjects,
                           total_tests=total_tests,
                           class_student_counts=class_student_counts,
                           subject_test_counts=subject_test_counts,
                           class_subject_map=class_subject_map,
                           school_name=school_name,
                           logo_path=logo_path)

@app.route('/admin/logout')
def admin_logout():
    """
    Handle admin logout.
    """
    session.pop('admin_logged_in', None)
    return redirect(url_for('admin_login'))

@app.route('/admin/exam_status')
@admin_required
def exam_status():
    """
    Render the exam status page showing all students and their progress.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT is_active FROM exam_control WHERE id=1")
    row = c.fetchone()
    conn.close()
    return jsonify({'is_active': bool(row['is_active']) if row else False})

@app.route('/admin/exam_has_started')
@admin_required
def exam_has_started():
    """Check if exam has ever been started (start_time is set)."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT start_time FROM exam_control WHERE id=1")
    row = c.fetchone()
    conn.close()
    has_started = row and row['start_time'] is not None
    return jsonify({'has_started': has_started})

@app.route('/admin/start_exam', methods=['POST'])
@admin_required
def start_exam():
    """
    API endpoint for admin to start the exam.
    Creates an exam_sessions record and updates exam_control for backward compatibility.
    """
    duration         = request.form.get('duration', 60, type=int)
    negative_marking = 1 if request.form.get('negative_marking') == 'yes' else 0
    negative_value   = request.form.get('negative_value', 0.33, type=float)
    cls              = request.form.get('class', '').strip()
    sec              = request.form.get('section', '').strip()
    subj             = request.form.get('subject', '').strip()
    tno              = request.form.get('test_no', '').strip()
    lab_label        = request.form.get('lab_label', '').strip() or 'Admin Console'
    lab_ip           = request.remote_addr

    conn = get_db()
    c = conn.cursor()
    start_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    c.execute("""
        INSERT INTO exam_sessions 
        (teacher_id, started_by_type, started_by_name, lab_ip, lab_label,
         class, section, subject, test_no, allotted_duration, start_time,
         status, negative_marking, negative_value)
        VALUES (NULL, 'admin', 'Admin', ?, ?, ?, ?, ?, ?, ?, ?, 'running', ?, ?)
    """, (lab_ip, lab_label, cls, sec, subj, tno, duration, start_time, negative_marking, negative_value))

    if tno:
        clean_cls = cls.replace('th','').replace('st','').replace('nd','').replace('rd','').strip() if cls else ''
        c.execute("""UPDATE test_papers SET is_active=1 WHERE test_no=?""", (tno,))
        if clean_cls and subj:
            c.execute("""UPDATE test_papers SET is_active=1 WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?)""",
                      (cls, f"%{clean_cls}%", subj))

    c.execute("""UPDATE exam_control
                 SET is_active=1, start_time=?, duration=?,
                     negative_marking=?, negative_value=?
                 WHERE id=1""",
              (start_time, duration, negative_marking, negative_value))
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'EXAM_START', 'exam_sessions', 1, request.remote_addr)
    return redirect(url_for('admin_dashboard'))

@app.route('/teacher/start_exam', methods=['POST'])
@teacher_required
def teacher_start_exam():
    """
    API endpoint for teacher to start an exam session.
    Isolates the session to this teacher's class/lab and prevents colliding with other teachers.
    """
    teacher_id       = session.get('teacher_id')
    duration         = request.form.get('duration', 60, type=int)
    negative_marking = 1 if request.form.get('negative_marking') == 'yes' else 0
    negative_value   = request.form.get('negative_value', 0.33, type=float)
    raw_assignment   = request.form.get('assignment', '').strip()
    cls              = request.form.get('class', '').strip()
    sec              = request.form.get('section', '').strip()
    subj             = request.form.get('subject', '').strip()
    tno              = request.form.get('test_no', '').strip()
    lab_label        = request.form.get('lab_label', '').strip()
    lab_ip           = request.remote_addr

    if raw_assignment and '|' in raw_assignment:
        parts = raw_assignment.split('|')
        if len(parts) >= 3:
            cls, sec, subj = parts[0].strip(), parts[1].strip(), parts[2].strip()
        elif len(parts) == 2:
            cls, subj = parts[0].strip(), parts[1].strip()

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT name FROM teachers WHERE id=?", (teacher_id,))
    t_row = c.fetchone()
    teacher_name = t_row['name'] if t_row else f"Teacher #{teacher_id}"

    if not lab_label:
        lab_label = f"Lab ({lab_ip})"

    assignments_to_start = []
    if cls and subj:
        assignments_to_start.append({'class': cls, 'section': sec, 'subject': subj})
    else:
        c.execute("SELECT class, section, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
        ta_rows = c.fetchall()
        if ta_rows:
            assignments_to_start = [dict(r) for r in ta_rows]
        else:
            assignments_to_start.append({'class': 'General', 'section': '', 'subject': 'General'})

    start_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    for a in assignments_to_start:
        c.execute("""
            INSERT INTO exam_sessions 
            (teacher_id, started_by_type, started_by_name, lab_ip, lab_label,
             class, section, subject, test_no, allotted_duration, start_time,
             status, negative_marking, negative_value)
            VALUES (?, 'teacher', ?, ?, ?, ?, ?, ?, ?, ?, ?, 'running', ?, ?)
        """, (teacher_id, teacher_name, lab_ip, lab_label,
              a['class'], a.get('section', ''), a['subject'], tno, duration, start_time,
              negative_marking, negative_value))

    if tno:
        clean_cls = cls.replace('th','').replace('st','').replace('nd','').replace('rd','').strip() if cls else ''
        c.execute("""UPDATE test_papers SET is_active=1 WHERE test_no=?""", (tno,))
        if clean_cls and subj:
            c.execute("""UPDATE test_papers SET is_active=1 WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?)""",
                      (cls, f"%{clean_cls}%", subj))

    # Also keep exam_control active for backward compatibility
    c.execute("""UPDATE exam_control
                 SET is_active=1, start_time=?, duration=?,
                     negative_marking=?, negative_value=?
                 WHERE id=1""",
              (start_time, duration, negative_marking, negative_value))

    conn.commit()
    conn.close()
    log_audit_event('teacher', teacher_id or 'teacher', 'EXAM_START', 'exam_sessions', 1, request.remote_addr)
    return redirect(url_for('teacher_dashboard'))

@app.route('/teacher/stop_exam')
@teacher_required
def teacher_stop_exam():
    """
    API endpoint for teacher to stop their own exam session(s).
    Ends ONLY this teacher's active sessions! Does NOT affect any other lab/teacher!
    """
    teacher_id = session.get('teacher_id')
    session_id = request.args.get('session_id', type=int)
    conn = get_db()
    c = conn.cursor()

    if session_id:
        c.execute("""UPDATE exam_sessions 
                     SET status='ended', end_time=CURRENT_TIMESTAMP, end_reason='ended_by_teacher'
                     WHERE id=? AND teacher_id=?""", (session_id, teacher_id))
    else:
        c.execute("""UPDATE exam_sessions 
                     SET status='ended', end_time=CURRENT_TIMESTAMP, end_reason='ended_by_teacher'
                     WHERE teacher_id=? AND status IN ('running', 'paused')""", (teacher_id,))

    # Check if ANY running/paused sessions remain in the whole system
    c.execute("SELECT COUNT(*) FROM exam_sessions WHERE status IN ('running', 'paused')")
    remaining_active = c.fetchone()[0]
    if remaining_active == 0:
        # Only set global is_active=0 if no other teacher/admin sessions are running
        c.execute("UPDATE exam_control SET is_active=0 WHERE id=1")

    conn.commit()
    conn.close()
    log_audit_event('teacher', teacher_id or 'teacher', 'EXAM_STOP', 'exam_sessions', session_id or 0, request.remote_addr)
    return redirect(url_for('teacher_dashboard'))

# ── Per-Session Exam Control API Routes (Admin & Teacher) ─────────────────
@app.route('/admin/exam_sessions')
@admin_required
def admin_exam_sessions():
    """
    Render dedicated Exam Control & Live Lab Sessions Panel for Admin.
    """
    return render_template('admin_exam_sessions.html')

@app.route('/admin/exam_sessions/data')
@admin_required
def admin_exam_sessions_data():
    """
    Return JSON list of all active (running/paused) and recent (last 24h) exam sessions.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM exam_sessions 
        WHERE status IN ('running', 'paused') 
           OR created_at >= datetime('now', '-24 hours')
        ORDER BY 
            CASE status WHEN 'running' THEN 1 WHEN 'paused' THEN 2 ELSE 3 END,
            id DESC
    """)
    rows = [dict(r) for r in c.fetchall()]
    sessions_data = []
    for s in rows:
        counts = get_session_student_counts(s, conn=conn)
        rem_secs = get_session_remaining_seconds(s)
        s_data = dict(s)
        s_data['remaining_seconds'] = rem_secs
        s_data['in_progress_count'] = counts['in_progress']
        s_data['submitted_count'] = counts['submitted']
        s_data['total_students'] = counts['total']
        sessions_data.append(s_data)
    conn.close()
    return jsonify({'status': 'success', 'sessions': sessions_data})

@app.route('/admin/exam_sessions/<int:session_id>/end', methods=['POST'])
@admin_required
def admin_end_exam_session(session_id):
    """
    Admin ends a specific running or paused exam session, optionally attaching a custom message.
    Force-submits all students scoped to that session.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM exam_sessions WHERE id=?", (session_id,))
    sess = c.fetchone()
    if not sess:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Session not found'}), 404

    sess = dict(sess)
    data = request.get_json(silent=True) or request.form
    end_msg = (data.get('message') or '').strip() or 'Exam ended by Administrator.'

    c.execute("""UPDATE exam_sessions 
                 SET status='ended', end_time=CURRENT_TIMESTAMP, end_reason='ended_by_admin', end_message=?
                 WHERE id=?""", (end_msg, session_id))

    # Force submit all in-progress students in this session scope
    cls = sess['class']
    clean_cls = str(cls).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    sec = sess['section']
    subj = sess['subject']
    tno = sess['test_no']
    where = "status='In Progress' AND (class=? OR class=? OR ? LIKE '%' || class || '%') AND LOWER(subject)=LOWER(?)"
    params = [cls, clean_cls, cls, subj]
    if sec:
        where += " AND LOWER(section)=LOWER(?)"
        params.append(sec)
    if tno:
        where += " AND (test_no=? OR test_no='' OR test_no IS NULL)"
        params.append(tno)
    c.execute(f"SELECT student_id FROM students WHERE {where}", params)
    ended_sids = [r['student_id'] for r in c.fetchall()]
    c.execute(f"UPDATE students SET status='Submitted' WHERE {where}", params)
    if ended_sids:
        process_coding_auto_submissions(ended_sids, conn=conn)

    # Check if other running/paused sessions remain in system
    c.execute("SELECT COUNT(*) FROM exam_sessions WHERE status IN ('running', 'paused')")
    if c.fetchone()[0] == 0:
        c.execute("UPDATE exam_control SET is_active=0 WHERE id=1")

    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'EXAM_SESSION_END', 'exam_sessions', session_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': f'Session #{session_id} ended successfully.'})

@app.route('/admin/exam_sessions/<int:session_id>/extend', methods=['POST'])
@admin_required
def admin_extend_exam_session(session_id):
    """
    Adjust allotted duration for a session (+ or - N minutes).
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM exam_sessions WHERE id=?", (session_id,))
    sess = c.fetchone()
    if not sess:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Session not found'}), 404

    data = request.get_json(silent=True) or request.form
    minutes = int(data.get('minutes', 5))
    curr_dur = int(sess['allotted_duration'] or 60)
    new_dur = max(1, curr_dur + minutes)

    c.execute("UPDATE exam_sessions SET allotted_duration=? WHERE id=?", (new_dur, session_id))
    conn.commit()

    c.execute("SELECT * FROM exam_sessions WHERE id=?", (session_id,))
    updated_sess = dict(c.fetchone())
    conn.close()

    rem = get_session_remaining_seconds(updated_sess)
    log_audit_event('admin', 'admin', 'EXAM_SESSION_EXTEND', 'exam_sessions', session_id, request.remote_addr)
    return jsonify({'status': 'success', 'new_duration': new_dur, 'remaining_seconds': rem})

@app.route('/admin/exam_sessions/<int:session_id>/pause', methods=['POST'])
@admin_required
def admin_pause_exam_session(session_id):
    """
    Pause a running session and freeze countdown.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM exam_sessions WHERE id=?", (session_id,))
    sess = c.fetchone()
    if not sess:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Session not found'}), 404

    if sess['status'] == 'running':
        now_str = datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
        c.execute("UPDATE exam_sessions SET status='paused', pause_started_at=? WHERE id=?", (now_str, session_id))
        conn.commit()

    conn.close()
    log_audit_event('admin', 'admin', 'EXAM_SESSION_PAUSE', 'exam_sessions', session_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': f'Session #{session_id} paused.'})

@app.route('/admin/exam_sessions/<int:session_id>/resume', methods=['POST'])
@admin_required
def admin_resume_exam_session(session_id):
    """
    Resume a paused session, accumulating paused duration.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM exam_sessions WHERE id=?", (session_id,))
    sess = c.fetchone()
    if not sess:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Session not found'}), 404

    if sess['status'] == 'paused' and sess['pause_started_at']:
        now = datetime.datetime.utcnow()
        p_start = parse_session_dt(sess['pause_started_at'])
        p_secs = max(0, int((now - p_start).total_seconds()))
        curr_paused = int(sess['total_paused_secs'] or 0)
        new_paused = curr_paused + p_secs
        c.execute("""UPDATE exam_sessions 
                     SET status='running', pause_started_at=NULL, total_paused_secs=?
                     WHERE id=?""", (new_paused, session_id))
        conn.commit()

    conn.close()
    log_audit_event('admin', 'admin', 'EXAM_SESSION_RESUME', 'exam_sessions', session_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': f'Session #{session_id} resumed.'})

@app.route('/admin/exam_sessions/<int:session_id>/label', methods=['POST', 'PATCH'])
@admin_required
def admin_label_exam_session(session_id):
    """
    Update friendly lab label (e.g. 'Lab 204 - Computer Science') for a session.
    """
    data = request.get_json(silent=True) or request.form
    label = (data.get('label') or '').strip()
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE exam_sessions SET lab_label=? WHERE id=?", (label, session_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'lab_label': label})

@app.route('/teacher/exam_sessions/data')
@teacher_required
def teacher_exam_sessions_data():
    """
    Return JSON list of this teacher's active and recent exam sessions.
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT * FROM exam_sessions 
        WHERE teacher_id=? 
          AND (status IN ('running', 'paused') OR created_at >= datetime('now', '-24 hours'))
        ORDER BY 
            CASE status WHEN 'running' THEN 1 WHEN 'paused' THEN 2 ELSE 3 END,
            id DESC
    """, (teacher_id,))
    rows = [dict(r) for r in c.fetchall()]
    sessions_data = []
    for s in rows:
        counts = get_session_student_counts(s, conn=conn)
        rem_secs = get_session_remaining_seconds(s)
        s_data = dict(s)
        s_data['remaining_seconds'] = rem_secs
        s_data['in_progress_count'] = counts['in_progress']
        s_data['submitted_count'] = counts['submitted']
        s_data['total_students'] = counts['total']
        sessions_data.append(s_data)
    conn.close()
    return jsonify({'status': 'success', 'sessions': sessions_data})

@app.route('/teacher/exam_sessions/<int:session_id>/end', methods=['POST'])
@teacher_required
def teacher_end_specific_exam_session(session_id):
    """
    Teacher ends their own specific exam session.
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM exam_sessions WHERE id=? AND teacher_id=?", (session_id, teacher_id))
    sess = c.fetchone()
    if not sess:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Session not found or unauthorized'}), 404

    sess = dict(sess)
    data = request.get_json(silent=True) or request.form
    end_msg = (data.get('message') or '').strip() or 'Exam ended by Teacher.'
    c.execute("""UPDATE exam_sessions 
                 SET status='ended', end_time=CURRENT_TIMESTAMP, end_reason='ended_by_teacher', end_message=?
                 WHERE id=?""", (end_msg, session_id))

    # Force submit students for this session
    cls = sess['class']
    clean_cls = str(cls).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    sec = sess['section']
    subj = sess['subject']
    tno = sess['test_no']
    where = "status='In Progress' AND (class=? OR class=? OR ? LIKE '%' || class || '%') AND LOWER(subject)=LOWER(?)"
    params = [cls, clean_cls, cls, subj]
    if sec:
        where += " AND LOWER(section)=LOWER(?)"
        params.append(sec)
    if tno:
        where += " AND (test_no=? OR test_no='' OR test_no IS NULL)"
        params.append(tno)
    c.execute(f"SELECT student_id FROM students WHERE {where}", params)
    ended_sids = [r['student_id'] for r in c.fetchall()]
    c.execute(f"UPDATE students SET status='Submitted' WHERE {where}", params)
    if ended_sids:
        process_coding_auto_submissions(ended_sids, conn=conn)

    # Check if other sessions are running before touching global is_active
    c.execute("SELECT COUNT(*) FROM exam_sessions WHERE status IN ('running', 'paused')")
    if c.fetchone()[0] == 0:
        c.execute("UPDATE exam_control SET is_active=0 WHERE id=1")

    conn.commit()
    conn.close()
    log_audit_event('teacher', teacher_id or 'teacher', 'EXAM_SESSION_END', 'exam_sessions', session_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': f'Session #{session_id} ended.'})


@app.route('/admin/exam_settings')
@admin_required
def admin_exam_settings():
    """Return current exam control settings as JSON — used by admin dashboard."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT duration, negative_marking, negative_value FROM exam_control WHERE id=1")
    row = c.fetchone()
    conn.close()
    if row:
        return jsonify({'duration': row['duration'],
                        'negative_marking': bool(row['negative_marking']),
                        'negative_value': row['negative_value']})
    return jsonify({'duration': 60, 'negative_marking': False, 'negative_value': 0.33})

@app.route('/teacher/exam_settings')
@teacher_required
def teacher_exam_settings():
    """Return current exam control settings as JSON — used by teacher dashboard."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT is_active, duration, negative_marking, negative_value FROM exam_control WHERE id=1")
    row = c.fetchone()
    conn.close()
    if row:
        return jsonify({'is_active': bool(row['is_active']),
                        'duration': row['duration'],
                        'negative_marking': bool(row['negative_marking']),
                        'negative_value': row['negative_value']})
    return jsonify({'is_active': False, 'duration': 60, 'negative_marking': False, 'negative_value': 0.33})

# ═══════════════════════════════════════════════════════════════
# TEST SCHEDULING ROUTES & APIS
# ═══════════════════════════════════════════════════════════════

@app.route('/admin/schedule_test')
@admin_required
def admin_schedule_test():
    update_scheduled_tests_status()
    class_subject_map = _get_class_subject_map()
    return render_template('scheduled_tests.html', is_admin=True, class_subject_map=class_subject_map)

@app.route('/teacher/schedule_test')
@teacher_required
def teacher_schedule_test():
    update_scheduled_tests_status()
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, section, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]
    conn.close()
    class_subject_map = _get_class_subject_map()
    return render_template('scheduled_tests.html', is_admin=False, assignments=assignments, class_subject_map=class_subject_map)

@app.route('/api/scheduled_test/create', methods=['POST'])
def create_scheduled_test():
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401

    class_ = request.form.get('class', '').strip()
    section = request.form.get('section', '').strip()
    subject = request.form.get('subject', '').strip()
    test_no = request.form.get('test_no', '').strip()
    scheduled_date = request.form.get('scheduled_date', '').strip()
    start_time_raw = request.form.get('start_time', '').strip()
    end_time_raw = request.form.get('end_time', '').strip()
    duration_minutes = int(request.form.get('duration_minutes', 30) or 30)

    if not class_ or not subject or not test_no or not scheduled_date or not start_time_raw or not end_time_raw:
        return jsonify({'status': 'error', 'message': 'Class, Subject, Test No, Date, Start Time, and End Time are required'}), 400

    full_start_str = f"{scheduled_date} {start_time_raw}:00" if len(start_time_raw) == 5 else f"{scheduled_date} {start_time_raw}"
    full_end_str = f"{scheduled_date} {end_time_raw}:00" if len(end_time_raw) == 5 else f"{scheduled_date} {end_time_raw}"

    try:
        st_dt = datetime.datetime.strptime(full_start_str, '%Y-%m-%d %H:%M:%S')
        end_dt = datetime.datetime.strptime(full_end_str, '%Y-%m-%d %H:%M:%S')
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Invalid Date or Time format: {e}'}), 400

    if end_dt <= st_dt:
        return jsonify({'status': 'error', 'message': 'End time must be after Start time'}), 400

    now = datetime.datetime.now()
    if st_dt <= now <= end_dt:
        init_status = 'active'
    elif now > end_dt:
        init_status = 'expired'
    else:
        init_status = 'scheduled'

    created_by = 'admin' if session.get('admin_logged_in') else str(session.get('teacher_id', 'teacher'))

    conn = get_db()
    c = conn.cursor()

    # Check if a schedule already exists for this Class, Section, Subject, and Test No
    c.execute("""SELECT id FROM scheduled_tests
                 WHERE class=? AND section=? AND LOWER(subject)=LOWER(?) AND test_no=?""",
              (class_, section, subject, test_no))
    existing = c.fetchone()

    if existing:
        # UPDATE existing row directly — no new row added to the list!
        c.execute("""UPDATE scheduled_tests
                     SET scheduled_date=?, start_time=?, end_time=?, duration_minutes=?, status=?, created_by=?
                     WHERE id=?""",
                  (scheduled_date, full_start_str, full_end_str, duration_minutes, init_status, created_by, existing['id']))
        conn.commit()
        conn.close()
        msg = f'Test schedule updated successfully ({init_status.capitalize()})'
    else:
        # Insert new row if no prior schedule exists for this test
        c.execute("""INSERT INTO scheduled_tests
                     (class, section, subject, test_no, scheduled_date, start_time, end_time, duration_minutes, status, created_by)
                     VALUES (?,?,?,?,?,?,?,?,?,?)""",
                  (class_, section, subject, test_no, scheduled_date, full_start_str, full_end_str, duration_minutes, init_status, created_by))
        conn.commit()
        conn.close()
        msg = f'Test scheduled successfully ({init_status.capitalize()})'

    u_type = 'admin' if session.get('admin_logged_in') else 'teacher'
    u_id = session.get('teacher_id') if session.get('teacher_logged_in') else 'admin'
    log_audit_event(u_type, u_id, 'TEST_SCHEDULE_SAVE', 'scheduled_tests', f"{class_}_{subject}_{test_no}", request.remote_addr)

    return jsonify({'status': 'success', 'message': msg})

@app.route('/api/scheduled_tests', methods=['GET'])
def get_scheduled_tests():
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401

    update_scheduled_tests_status()

    class_ = request.args.get('class', '').strip()
    section = request.args.get('section', '').strip()
    date_filter = request.args.get('date', '').strip()
    status_filter = request.args.get('status', '').strip()

    conn = get_db()
    c = conn.cursor()
    query = "SELECT * FROM scheduled_tests WHERE 1=1"
    params = []

    if class_:
        query += " AND class=?"
        params.append(class_)
    if section:
        query += " AND section=?"
        params.append(section)
    if date_filter:
        query += " AND scheduled_date=?"
        params.append(date_filter)
    if status_filter:
        query += " AND status=?"
        params.append(status_filter)

    query += " ORDER BY id DESC"
    c.execute(query, params)
    rows = [dict(r) for r in c.fetchall()]
    conn.close()

    return jsonify({'status': 'success', 'scheduled_tests': rows})

@app.route('/api/scheduled_test/<int:sid>/cancel', methods=['POST'])
def cancel_scheduled_test(sid):
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE scheduled_tests SET status='cancelled' WHERE id=?", (sid,))
    conn.commit()
    conn.close()
    u_type = 'admin' if session.get('admin_logged_in') else 'teacher'
    u_id = session.get('teacher_id') if session.get('teacher_logged_in') else 'admin'
    log_audit_event(u_type, u_id, 'TEST_SCHEDULE_CANCEL', 'scheduled_tests', sid, request.remote_addr)
    return jsonify({'status': 'success', 'message': 'Scheduled test cancelled'})

@app.route('/api/scheduled_test/<int:sid>/reopen', methods=['POST'])
def reopen_scheduled_test(sid):
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401

    scheduled_date = request.form.get('scheduled_date', '').strip()
    start_time_raw = request.form.get('start_time', '').strip()
    end_time_raw = request.form.get('end_time', '').strip()
    duration_minutes = int(request.form.get('duration_minutes', 30) or 30)

    if not scheduled_date or not start_time_raw or not end_time_raw:
        return jsonify({'status': 'error', 'message': 'Date, Start Time, and End Time are required'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM scheduled_tests WHERE id=?", (sid,))
    orig = c.fetchone()
    if not orig:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Schedule entry not found'}), 404

    full_start_str = f"{scheduled_date} {start_time_raw}:00" if len(start_time_raw) == 5 else f"{scheduled_date} {start_time_raw}"
    full_end_str = f"{scheduled_date} {end_time_raw}:00" if len(end_time_raw) == 5 else f"{scheduled_date} {end_time_raw}"

    try:
        st_dt = datetime.datetime.strptime(full_start_str, '%Y-%m-%d %H:%M:%S')
        end_dt = datetime.datetime.strptime(full_end_str, '%Y-%m-%d %H:%M:%S')
    except Exception as e:
        conn.close()
        return jsonify({'status': 'error', 'message': f'Invalid Date or Time format: {e}'}), 400

    if end_dt <= st_dt:
        conn.close()
        return jsonify({'status': 'error', 'message': 'End time must be after Start time'}), 400

    now = datetime.datetime.now()
    if st_dt <= now <= end_dt:
        init_status = 'active'
    elif now > end_dt:
        init_status = 'expired'
    else:
        init_status = 'scheduled'

    created_by = 'admin' if session.get('admin_logged_in') else str(session.get('teacher_id', 'teacher'))

    # Update existing schedule row directly — status changes to active/scheduled/expired without creating duplicate row
    c.execute("""UPDATE scheduled_tests
                 SET scheduled_date=?, start_time=?, end_time=?, duration_minutes=?, status=?, created_by=?
                 WHERE id=?""",
              (scheduled_date, full_start_str, full_end_str, duration_minutes, init_status, created_by, sid))
    conn.commit()
    conn.close()

    u_type = 'admin' if session.get('admin_logged_in') else 'teacher'
    u_id = session.get('teacher_id') if session.get('teacher_logged_in') else 'admin'
    log_audit_event(u_type, u_id, 'TEST_SCHEDULE_RESCHEDULE', 'scheduled_tests', sid, request.remote_addr)

    return jsonify({'status': 'success', 'message': f'Test rescheduled successfully ({init_status.capitalize()})'})

@app.route('/api/scheduled_test/check_student', methods=['GET'])
def check_student_scheduled_test():
    student_id = session.get('student_id')
    if not student_id:
        return jsonify({'status': 'error', 'message': 'Not logged in'}), 401

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT class, section, subject, test_no, status FROM students WHERE student_id=?", (student_id,))
    st = c.fetchone()
    conn.close()

    if not st:
        return jsonify({'status': 'error', 'message': 'Student not found'}), 404

    sch_res = get_student_schedule_status(st['class'], st['section'], st['subject'], st['test_no'], student_id)
    return jsonify({'status': 'success', 'schedule_info': sch_res})

@app.route('/api/available_test_numbers', methods=['GET'])
def get_available_test_numbers():
    class_ = request.args.get('class', '').strip()
    subject = request.args.get('subject', '').strip()
    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()

    conn = get_db()
    c = conn.cursor()
    if class_ and subject:
        c.execute("""SELECT DISTINCT test_no FROM test_papers 
                     WHERE (class=? OR class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND is_active=1 AND test_no IS NOT NULL AND trim(test_no) != ''
                     UNION
                     SELECT DISTINCT test_no FROM questions 
                     WHERE (class=? OR class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no IS NOT NULL AND trim(test_no) != ''
                       AND test_no NOT IN (
                           SELECT test_no FROM test_papers 
                           WHERE (class=? OR class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND is_active=0 AND test_no IS NOT NULL
                       )
                     ORDER BY test_no""",
                  (class_, clean_cls, f"%{clean_cls}%", subject, 
                   class_, clean_cls, f"%{clean_cls}%", subject,
                   class_, clean_cls, f"%{clean_cls}%", subject))
    elif class_:
        c.execute("""SELECT DISTINCT test_no FROM test_papers 
                     WHERE (class=? OR class=? OR class LIKE ?) AND is_active=1 AND test_no IS NOT NULL AND trim(test_no) != ''
                     UNION
                     SELECT DISTINCT test_no FROM questions 
                     WHERE (class=? OR class=? OR class LIKE ?) AND test_no IS NOT NULL AND trim(test_no) != ''
                       AND test_no NOT IN (
                           SELECT test_no FROM test_papers 
                           WHERE (class=? OR class=? OR class LIKE ?) AND is_active=0 AND test_no IS NOT NULL
                       )
                     ORDER BY test_no""",
                  (class_, clean_cls, f"%{clean_cls}%", 
                   class_, clean_cls, f"%{clean_cls}%",
                   class_, clean_cls, f"%{clean_cls}%"))
    elif subject:
        c.execute("""SELECT DISTINCT test_no FROM test_papers 
                     WHERE LOWER(subject)=LOWER(?) AND is_active=1 AND test_no IS NOT NULL AND trim(test_no) != ''
                     UNION
                     SELECT DISTINCT test_no FROM questions 
                     WHERE LOWER(subject)=LOWER(?) AND test_no IS NOT NULL AND trim(test_no) != ''
                       AND test_no NOT IN (
                           SELECT test_no FROM test_papers 
                           WHERE LOWER(subject)=LOWER(?) AND is_active=0 AND test_no IS NOT NULL
                       )
                     ORDER BY test_no""", (subject, subject, subject))
    else:
        c.execute("""SELECT DISTINCT test_no FROM test_papers WHERE is_active=1 AND test_no IS NOT NULL AND trim(test_no) != ''
                     UNION
                     SELECT DISTINCT test_no FROM questions WHERE test_no IS NOT NULL AND trim(test_no) != ''
                       AND test_no NOT IN (
                           SELECT test_no FROM test_papers WHERE is_active=0 AND test_no IS NOT NULL
                       )
                     ORDER BY test_no""")
    rows = [r[0] for r in c.fetchall() if r[0]]
    conn.close()
    return jsonify({'status': 'success', 'test_numbers': rows})

@app.route('/admin/student/force_submit/<student_id>', methods=['POST'])
@admin_required
def admin_force_submit_student(student_id):
    """Feature #3: Force-submit an individual student from monitoring page."""
    conn = get_db()
    c = conn.cursor()

    # Only act if student is In Progress
    c.execute("SELECT name, class, section, subject, status, test_no FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student not found'}), 404
    if student['status'] not in ('In Progress', 'Not Started'):
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student already submitted'}), 400

    name    = student['name']
    class_  = student['class']
    subject = student['subject']
    test_no = (student['test_no'] if ('test_no' in student.keys() and student['test_no']) else '').strip()

    # Calculate score with negative marking
    c.execute("SELECT negative_marking, negative_value FROM exam_control WHERE id=1")
    ec = c.fetchone()
    neg_enabled = bool(ec['negative_marking']) if ec else False
    neg_value   = float(ec['negative_value'])  if ec else 0.33

    if test_no:
        c.execute("""SELECT q.correct_answer, q.option_a, q.option_b, q.option_c, q.option_d, r.selected_option
                     FROM questions q
                     LEFT JOIN responses r ON q.id = r.question_id AND r.student_id = ?
                     WHERE q.class = ? AND q.subject = ? AND (q.test_no = ? OR q.chapter = ?)""",
                  (student_id, class_, subject, test_no, test_no))
    else:
        c.execute("""SELECT q.correct_answer, q.option_a, q.option_b, q.option_c, q.option_d, r.selected_option
                     FROM questions q
                     LEFT JOIN responses r ON q.id = r.question_id AND r.student_id = ?
                     WHERE q.class = ? AND q.subject = ?""",
                  (student_id, class_, subject))
    rows  = c.fetchall()
    total = len(rows)
    raw = 0.0
    for r in rows:
        sel = (r['selected_option'] or '').strip().upper()
        cor_letter = _normalize_correct_answer(
            r['correct_answer'],
            r['option_a'],
            r['option_b'],
            r['option_c'],
            r['option_d']
        )
        if sel and cor_letter and sel == cor_letter:
            raw += 1.0
        elif sel and neg_enabled:
            raw -= neg_value
    score      = max(0.0, raw)
    percentage = round((score / total * 100), 2) if total > 0 else 0.0

    # Save result
    _save_or_update_result(c, student_id, name, class_, student['section'] if 'section' in student.keys() else '', subject, score, total, percentage, test_no)
    c.execute("UPDATE students SET status='Submitted' WHERE student_id=?", (student_id,))
    process_coding_auto_submissions([student_id], conn=conn)
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': f'{name} submitted. Score: {round(score,2)}/{total}'})

@app.route('/admin/stop_exam')
@admin_required
def stop_exam():
    """
    API endpoint for admin to stop all exams globally (Master Kill Switch).
    Marks all running/paused sessions as ended and sets exam_control.is_active=0.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE exam_control SET is_active=0 WHERE id=1")
    c.execute("""UPDATE exam_sessions 
                 SET status='ended', end_time=CURRENT_TIMESTAMP, end_reason='ended_by_admin', end_message='Exam stopped globally by Administrator.'
                 WHERE status IN ('running', 'paused')""")
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'EXAM_STOP_GLOBAL', 'exam_control', 1, request.remote_addr)
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/settings', methods=['GET', 'POST'])
@admin_required
def admin_settings():
    """
    Render and handle admin settings page (school name, logo, etc).
    """
    if request.method == 'POST':
        for field in ['school_name', 'school_address', 'academic_session']:
            val = request.form.get(field, '').strip()
            if val:
                set_setting(field, val)

        logo_file = request.files.get('logo')
        if logo_file and logo_file.filename != '' and allowed_image_file(logo_file.filename):
            ext = logo_file.filename.rsplit('.', 1)[1].lower()
            filename = f"school_logo.{ext}"
            filepath = os.path.join(app.config['BRANDING_FOLDER'], filename)
            for old_ext in ALLOWED_IMAGE_EXTENSIONS:
                old_path = os.path.join(app.config['BRANDING_FOLDER'], f"school_logo.{old_ext}")
                if os.path.exists(old_path) and old_path != filepath:
                    os.remove(old_path)
            logo_file.save(filepath)
            set_setting('logo_path', os.path.join('uploads', 'branding', filename).replace('\\', '/'))

        return redirect(url_for('admin_settings'))

    school_name      = get_setting('school_name', 'RRB Group of Schools')
    school_address   = get_setting('school_address', '')
    academic_session = get_setting('academic_session', '')
    logo_path        = get_setting('logo_path', '')
    return render_template('settings.html',
                           school_name=school_name,
                           school_address=school_address,
                           academic_session=academic_session,
                           logo_path=logo_path)

@app.route('/admin/generate_question_paper')
@app.route('/teacher/generate_question_paper')
def generate_question_paper():
    """
    Render page to generate HTML Print Preview / PDF of the question paper.
    """
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return redirect(url_for('teacher_login'))

    class_ = request.args.get('class', '').strip()
    section = request.args.get('section', '').strip()
    subject = request.args.get('subject', '').strip()
    test_no = request.args.get('test_no', '').strip()

    if not class_ or not subject:
        return "Class and subject required", 400

    conn = get_db()
    c = conn.cursor()

    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()

    # Query matching questions using multi-tier fallback
    c.execute("""SELECT id, question, option_a, option_b, option_c, option_d, correct_answer, image_path, chapter, test_no, board,
                        CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                 FROM questions
                 WHERE (class=? OR class=? OR class LIKE ?)
                   AND (? = '' OR section=? OR section LIKE ?)
                   AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                   AND (? = '' OR test_no=? OR chapter=? OR test_no LIKE ? OR chapter LIKE ?)
                 ORDER BY id""",
              (class_, clean_cls, f"%{clean_cls}%",
               section, section, f"%{section}%",
               subject, f"%{subject}%",
               test_no, test_no, test_no, f"%{test_no}%", f"%{test_no}%"))
    questions = [dict(row) for row in c.fetchall()]

    if not questions:
        c.execute("""SELECT id, question, option_a, option_b, option_c, option_d, correct_answer, image_path, chapter, test_no, board,
                            CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                     FROM questions
                     WHERE (LOWER(class)=LOWER(?) OR LOWER(class)=LOWER(?) OR class LIKE ?)
                       AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                     ORDER BY id""",
                  (class_, clean_cls, f"%{clean_cls}%", subject, f"%{subject}%"))
        questions = [dict(row) for row in c.fetchall()]

    conn.close()

    for q in questions:
        q_type = str(q.get('question_type') or '').strip()
        if not q_type or q_type.upper() in ['MCQ', 'CHOICE', 'OBJECTIVE']:
            q['question_type'] = 'MCQ'

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')
    logo_base64 = None
    if logo_path:
        abs_path = os.path.join(app.static_folder, logo_path)
        if os.path.exists(abs_path):
            with open(abs_path, "rb") as f:
                logo_bytes = f.read()
                ext = os.path.splitext(abs_path)[1].lower()
                mime = "image/png" if ext == ".png" else "image/jpeg"
                logo_base64 = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode('utf-8')}"

    # Determine display chapter name: prefer explicit query param or real question chapter over literal test_no
    req_chapter = request.args.get('chapter', '').strip()
    if req_chapter:
        chapter_display = req_chapter
    else:
        unique_ch = []
        for q in questions:
            ch_val = str(q.get('chapter') or '').strip()
            if ch_val and ch_val.lower() != test_no.lower() and ch_val not in unique_ch:
                unique_ch.append(ch_val)
        if unique_ch:
            chapter_display = ", ".join(unique_ch)
        else:
            chapter_display = 'All Chapters'

    req_board = request.args.get('board', '').strip()
    board_display = req_board or (questions[0].get('board') if questions and questions[0].get('board') else 'CBSE')

    total_marks = len(questions)
    display_class = f"{class_}{section}" if section else class_
    rendered_html = render_template(
        'question_paper_template_v2.html',
        school_name=school_name,
        logo_base64=logo_base64,
        class_name=display_class,
        subject=subject,
        chapter=chapter_display,
        test_no=test_no or 'All',
        board=board_display,
        questions=questions,
        total_marks=total_marks,
        date=datetime.datetime.now().strftime('%d/%m/%Y')
    )

    if request.args.get('format') == 'download_pdf':
        pdf = HTML(string=rendered_html, base_url=request.base_url).write_pdf()
        response = make_response(pdf)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=Question_Paper_{display_class}_{subject}.pdf'
        return response

    return rendered_html

@app.route('/admin/questions')
@admin_required
def questions():
    """
    Render the admin questions management page.
    """
    return render_template('questions.html', role='admin', api_base='/admin')

@app.route('/admin/questions/data')
@admin_required
def questions_data():
    """
    API endpoint to fetch questions data with filtering and pagination.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, class, section, subject, question, option_a, option_b, option_c, option_d, correct_answer, image_path, test_no, chapter FROM questions ORDER BY class, subject, id")
    questions = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(questions)

@app.route('/admin/question/add', methods=['POST'])
@admin_required
def add_question():
    """
    API endpoint to add a new question to the database.
    """
    if request.is_json:
        data = request.get_json()
        class_ = data['class']
        subject = data['subject']
        question = data['question']
        opt_a = data['option_a']
        opt_b = data['option_b']
        opt_c = data['option_c']
        opt_d = data['option_d']
        correct = data['correct_answer']
        image_file = None
    else:
        class_ = request.form.get('class')
        section = request.form.get('section', '').strip()
        subject = request.form.get('subject')
        test_no = request.form.get('test_no', '').strip()
        question = request.form.get('question')
        opt_a = request.form.get('option_a')
        opt_b = request.form.get('option_b')
        opt_c = request.form.get('option_c')
        opt_d = request.form.get('option_d')
        correct = request.form.get('correct_answer')
        image_file = request.files.get('image')

    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT INTO questions (class, section, subject, test_no, chapter, question, option_a, option_b, option_c, option_d, correct_answer) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
              (class_, section, subject, test_no, test_no, question, opt_a, opt_b, opt_c, opt_d, correct))
    qid = c.lastrowid

    image_path = None
    if image_file and image_file.filename != '':
        image_path = save_question_image(image_file, class_, subject, qid)
        if image_path:
            c.execute("UPDATE questions SET image_path=? WHERE id=?", (image_path, qid))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'id': qid})

@app.route('/admin/question/update/<int:qid>', methods=['POST'])
@admin_required
def update_question(qid):
    """
    API endpoint to update an existing question.
    """
    class_ = request.form.get('class')
    section = request.form.get('section', '').strip()
    subject = request.form.get('subject')
    test_no = request.form.get('test_no', '').strip()
    question = request.form.get('question')
    opt_a = request.form.get('option_a')
    opt_b = request.form.get('option_b')
    opt_c = request.form.get('option_c')
    opt_d = request.form.get('option_d')
    correct = request.form.get('correct_answer')
    image_file = request.files.get('image')
    remove_image = request.form.get('remove_image') == 'true'

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT image_path FROM questions WHERE id=?", (qid,))
    old_image = c.fetchone()
    old_path = old_image['image_path'] if old_image else None

    c.execute('''UPDATE questions 
                 SET class=?, section=?, subject=?, test_no=?, chapter=?, question=?, option_a=?, option_b=?, option_c=?, option_d=?, correct_answer=?
                 WHERE id=?''',
              (class_, section, subject, test_no, test_no, question, opt_a, opt_b, opt_c, opt_d, correct, qid))

    image_path = old_path
    if remove_image:
        if old_path and os.path.exists(os.path.join('static', old_path)):
            try:
                os.remove(os.path.join('static', old_path))
            except:
                pass
        image_path = None
        c.execute("UPDATE questions SET image_path=NULL WHERE id=?", (qid,))
    elif image_file and image_file.filename != '':
        if old_path and os.path.exists(os.path.join('static', old_path)):
            try:
                os.remove(os.path.join('static', old_path))
            except:
                pass
        new_path = save_question_image(image_file, class_, subject, qid)
        if new_path:
            image_path = new_path
            c.execute("UPDATE questions SET image_path=? WHERE id=?", (image_path, qid))

    conn.commit()
    conn.close()
    return jsonify({'status': 'updated'})

@app.route('/admin/question/delete/<int:qid>', methods=['DELETE'])
@admin_required
def delete_question(qid):
    """
    API endpoint to delete a question by ID.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT image_path FROM questions WHERE id=?", (qid,))
    row = c.fetchone()
    if row and row['image_path']:
        img_path = os.path.join('static', row['image_path'])
        if os.path.exists(img_path):
            try:
                os.remove(img_path)
            except:
                pass
    c.execute("DELETE FROM questions WHERE id=?", (qid,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'deleted'})

@app.route('/admin/questions/delete_by_class_subject', methods=['POST'])
@admin_required
def delete_questions_by_class_subject():
    """
    API endpoint to delete all questions for a specific class and subject.
    """
    data = request.get_json()
    class_ = data.get('class')
    subject = data.get('subject')

    if not class_ or not subject:
        return jsonify({'status': 'error', 'message': 'Class and subject required'}), 400

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT image_path FROM questions WHERE class=? AND subject=?", (class_, subject))
    images = c.fetchall()

    for img in images:
        if img['image_path']:
            img_path = os.path.join('static', img['image_path'])
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except:
                    pass

    c.execute("DELETE FROM questions WHERE class=? AND subject=?", (class_, subject))
    deleted_count = c.rowcount
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'PAPER_DELETE', 'questions', f"{class_}_{subject}", request.remote_addr)

    return jsonify({'status': 'success', 'deleted': deleted_count})

# -------------------------------
# Teacher Questions Management Routes
# -------------------------------
@app.route('/teacher/questions')
@teacher_required
def teacher_questions():
    """
    Render the teacher questions management page.
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('questions.html', role='teacher', api_base='/teacher', assignments=assignments)

@app.route('/teacher/questions/data')
@teacher_required
def teacher_questions_data():
    """
    API endpoint to fetch questions data for teacher's allotted classes and subjects.
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = c.fetchall()

    if assignments:
        where_clauses = []
        params = []
        for a in assignments:
            cls = str(a['class']).strip()
            clean_cls = cls.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            subj = str(a['subject']).strip()
            where_clauses.append("((class=? OR class=? OR class LIKE ?) AND (LOWER(subject)=LOWER(?) OR subject LIKE ?))")
            params.extend([cls, clean_cls, f"%{clean_cls}%", subj, f"%{subj}%"])
        sql = "SELECT id, class, section, subject, question, option_a, option_b, option_c, option_d, correct_answer, image_path, test_no, chapter FROM questions WHERE " + " OR ".join(where_clauses) + " ORDER BY class, subject, id"
        c.execute(sql, tuple(params))
    else:
        c.execute("SELECT id, class, section, subject, question, option_a, option_b, option_c, option_d, correct_answer, image_path, test_no, chapter FROM questions ORDER BY class, subject, id")

    questions = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(questions)

@app.route('/teacher/question/add', methods=['POST'])
@teacher_required
def teacher_add_question():
    """
    API endpoint for teachers to add a new question.
    """
    if request.is_json:
        data = request.get_json()
        class_ = data.get('class', '')
        section = data.get('section', '').strip()
        subject = data.get('subject', '')
        test_no = data.get('test_no', '').strip()
        question = data.get('question', '')
        opt_a = data.get('option_a', '')
        opt_b = data.get('option_b', '')
        opt_c = data.get('option_c', '')
        opt_d = data.get('option_d', '')
        correct = data.get('correct_answer', '')
        image_file = None
    else:
        class_ = request.form.get('class', '')
        section = request.form.get('section', '').strip()
        subject = request.form.get('subject', '')
        test_no = request.form.get('test_no', '').strip()
        question = request.form.get('question', '')
        opt_a = request.form.get('option_a', '')
        opt_b = request.form.get('option_b', '')
        opt_c = request.form.get('option_c', '')
        opt_d = request.form.get('option_d', '')
        correct = request.form.get('correct_answer', '')
        image_file = request.files.get('image')

    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT INTO questions (class, section, subject, test_no, chapter, question, option_a, option_b, option_c, option_d, correct_answer) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
              (class_, section, subject, test_no, test_no, question, opt_a, opt_b, opt_c, opt_d, correct))
    qid = c.lastrowid

    image_path = None
    if image_file and image_file.filename != '':
        image_path = save_question_image(image_file, class_, subject, qid)
        if image_path:
            c.execute("UPDATE questions SET image_path=? WHERE id=?", (image_path, qid))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'id': qid})

@app.route('/teacher/question/update/<int:qid>', methods=['POST'])
@teacher_required
def teacher_update_question(qid):
    """
    API endpoint for teachers to update an existing question.
    """
    class_ = request.form.get('class', '')
    section = request.form.get('section', '').strip()
    subject = request.form.get('subject', '')
    test_no = request.form.get('test_no', '').strip()
    question = request.form.get('question', '')
    opt_a = request.form.get('option_a', '')
    opt_b = request.form.get('option_b', '')
    opt_c = request.form.get('option_c', '')
    opt_d = request.form.get('option_d', '')
    correct = request.form.get('correct_answer', '')
    image_file = request.files.get('image')
    remove_image = request.form.get('remove_image') == 'true'

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT image_path FROM questions WHERE id=?", (qid,))
    old_image = c.fetchone()
    old_path = old_image['image_path'] if old_image else None

    c.execute('''UPDATE questions 
                 SET class=?, section=?, subject=?, test_no=?, chapter=?, question=?, option_a=?, option_b=?, option_c=?, option_d=?, correct_answer=?
                 WHERE id=?''',
              (class_, section, subject, test_no, test_no, question, opt_a, opt_b, opt_c, opt_d, correct, qid))

    image_path = old_path
    if remove_image:
        if old_path and os.path.exists(os.path.join('static', old_path)):
            try:
                os.remove(os.path.join('static', old_path))
            except:
                pass
        image_path = None
        c.execute("UPDATE questions SET image_path=NULL WHERE id=?", (qid,))
    elif image_file and image_file.filename != '':
        if old_path and os.path.exists(os.path.join('static', old_path)):
            try:
                os.remove(os.path.join('static', old_path))
            except:
                pass
        new_path = save_question_image(image_file, class_, subject, qid)
        if new_path:
            image_path = new_path
            c.execute("UPDATE questions SET image_path=? WHERE id=?", (image_path, qid))

    conn.commit()
    conn.close()
    return jsonify({'status': 'updated'})

@app.route('/teacher/question/delete/<int:qid>', methods=['DELETE'])
@teacher_required
def teacher_delete_question(qid):
    """
    API endpoint for teachers to delete a question by ID.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT image_path FROM questions WHERE id=?", (qid,))
    row = c.fetchone()
    if row and row['image_path']:
        img_path = os.path.join('static', row['image_path'])
        if os.path.exists(img_path):
            try:
                os.remove(img_path)
            except:
                pass
    c.execute("DELETE FROM questions WHERE id=?", (qid,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'deleted'})

@app.route('/teacher/questions/delete_by_class_subject', methods=['POST'])
@teacher_required
def teacher_delete_questions_by_class_subject():
    """
    API endpoint for teachers to delete all questions for a specific class and subject.
    """
    data = request.get_json()
    class_ = data.get('class')
    subject = data.get('subject')

    if not class_ or not subject:
        return jsonify({'status': 'error', 'message': 'Class and subject required'}), 400

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT image_path FROM questions WHERE class=? AND subject=?", (class_, subject))
    images = c.fetchall()

    for img in images:
        if img['image_path']:
            img_path = os.path.join('static', img['image_path'])
            if os.path.exists(img_path):
                try:
                    os.remove(img_path)
                except:
                    pass

    c.execute("DELETE FROM questions WHERE class=? AND subject=?", (class_, subject))
    deleted_count = c.rowcount
    conn.commit()
    conn.close()
    teacher_id = str(session.get('teacher_id', 'teacher'))
    log_audit_event('teacher', teacher_id, 'PAPER_DELETE', 'questions', f"{class_}_{subject}", request.remote_addr)

    return jsonify({'status': 'success', 'deleted': deleted_count})

@app.route('/admin/upload_csv', methods=['POST'])
@admin_required
def upload_csv():
    """
    API endpoint to upload questions from a CSV file.
    """
    if 'csv_file' not in request.files:
        return redirect(request.url)
    file = request.files['csv_file']
    if file.filename == '':
        return redirect(request.url)
    if file and file.filename.endswith('.csv'):
        filename = secure_filename(file.filename)
        base = os.path.splitext(filename)[0]
        parts = base.split('_')
        # Parse class, subject, and optional test_no
        if len(parts) >= 3:
            class_ = parts[0].strip()
            subject = parts[1].strip()
            test_no = parts[2].strip()
        elif len(parts) == 2:
            class_ = parts[0].strip()
            subject = parts[1].strip()
            test_no = ''
        else:
            class_ = "Unknown"
            subject = base.strip()
            test_no = ''
        if not class_:
            class_ = "Unknown"

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        conn = get_db()
        c = conn.cursor()
        try:
            with open(filepath, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                header = next(reader)
                normalized = [h.strip().lower().replace(' ', '_') for h in header]

                try:
                    idx_q = normalized.index('question')
                    idx_a = normalized.index('option_a')
                    idx_b = normalized.index('option_b')
                    idx_c = normalized.index('option_c')
                    idx_d = normalized.index('option_d')
                    idx_correct = normalized.index('correct_answer')
                except ValueError as e:
                    return f"Missing required column: {e}. Headers found: {header}", 400

                inserted_count = 0
                for row in reader:
                    if not row or all(cell.strip() == '' for cell in row):
                        continue
                    if len(row) <= max(idx_q, idx_a, idx_b, idx_c, idx_d, idx_correct):
                        continue

                    question = row[idx_q].strip()
                    opt_a = row[idx_a].strip()
                    opt_b = row[idx_b].strip()
                    opt_c = row[idx_c].strip()
                    opt_d = row[idx_d].strip()
                    correct_value = row[idx_correct].strip()

                    correct_letter = None
                    if correct_value == opt_a:
                        correct_letter = 'A'
                    elif correct_value == opt_b:
                        correct_letter = 'B'
                    elif correct_value == opt_c:
                        correct_letter = 'C'
                    elif correct_value == opt_d:
                        correct_letter = 'D'
                    else:
                        if correct_value.upper() in ('A', 'B', 'C', 'D'):
                            correct_letter = correct_value.upper()
                        elif correct_value.lower() == 'option_a':
                            correct_letter = 'A'
                        elif correct_value.lower() == 'option_b':
                            correct_letter = 'B'
                        elif correct_value.lower() == 'option_c':
                            correct_letter = 'C'
                        elif correct_value.lower() == 'option_d':
                            correct_letter = 'D'
                        else:
                            print(f"Warning: Could not map correct_answer '{correct_value}' to any option. Skipping question.")
                            continue

                    if not question:
                        continue

                    # Insert with chapter (test_no)
                    c.execute("INSERT INTO questions (class, subject, question, option_a, option_b, option_c, option_d, correct_answer, chapter) VALUES (?,?,?,?,?,?,?,?,?)",
                              (class_, subject, question, opt_a, opt_b, opt_c, opt_d, correct_letter, test_no))
            conn.commit()
        except Exception as e:
            conn.rollback()
            return f"Error processing CSV: {e}", 500
        finally:
            conn.close()
            os.remove(filepath)
        return redirect(url_for('questions'))
    return "Invalid file", 400

@app.route('/admin/students')
@admin_required
def manage_students():
    """
    Render the admin students management page.
    """
    return render_template('manage_students.html')

@app.route('/admin/students/data')
@admin_required
def students_data():
    """
    API endpoint to fetch students data with advanced filtering:
    q (search name, admission_no, student_id), class, section, subject, status, test_no, date.
    Also returns system_name (instead of IP) and exam_started_at.
    """
    filter_q       = request.args.get('q', '').strip()
    filter_class   = request.args.get('class', '').strip()
    filter_section = request.args.get('section', '').strip()
    filter_subject = request.args.get('subject', '').strip()
    filter_status  = request.args.get('status', '').strip()
    filter_test_no = request.args.get('test_no', '').strip()
    filter_date    = request.args.get('date', '').strip()

    conn = get_db()
    c = conn.cursor()

    query = """SELECT student_id, name, class, section, subject,
                      ip, system_name, status, exam_started_at,
                      test_no, admission_no, dob, house, parents_name, address, picture
               FROM students WHERE 1=1"""
    params = []

    if filter_q:
        query += " AND (name LIKE ? OR admission_no LIKE ? OR student_id LIKE ?)"
        q_wild = f"%{filter_q}%"
        params.extend([q_wild, q_wild, q_wild])
    if filter_class:
        query += " AND class = ?"
        params.append(filter_class)
    if filter_section:
        query += " AND section = ?"
        params.append(filter_section)
    if filter_subject:
        query += " AND subject = ?"
        params.append(filter_subject)
    if filter_status:
        query += " AND status = ?"
        params.append(filter_status)
    if filter_test_no:
        query += " AND test_no LIKE ?"
        params.append(f"%{filter_test_no}%")
    if filter_date:
        query += " AND DATE(exam_started_at) = ?"
        params.append(filter_date)

    query += " ORDER BY class, section, name"
    c.execute(query, params)
    students = [dict(row) for row in c.fetchall()]

    # Fetch distinct classes and sections across all students for filter options
    c.execute("SELECT DISTINCT class FROM students WHERE class IS NOT NULL AND class != '' ORDER BY class")
    classes = [r['class'] for r in c.fetchall()]
    c.execute("SELECT DISTINCT section FROM students WHERE section IS NOT NULL AND section != '' ORDER BY section")
    sections = [r['section'] for r in c.fetchall()]
    c.execute("SELECT DISTINCT subject FROM students WHERE subject IS NOT NULL AND subject != '' ORDER BY subject")
    subjects = [r['subject'] for r in c.fetchall()]

    conn.close()
    return jsonify({
        'status': 'success',
        'students': students,
        'classes': classes,
        'sections': sections,
        'subjects': subjects,
        'total': len(students)
    })

@app.route('/admin/student/edit/<student_id>', methods=['POST'])
@admin_required
def admin_edit_student(student_id):
    """
    API endpoint for Admin to edit student information.
    """
    data = request.get_json(silent=True) or request.form.to_dict() or {}
    name = (data.get('name') or '').strip()
    admission_no = (data.get('admission_no') or '').strip()
    class_ = (data.get('class') or '').strip()
    section = (data.get('section') or '').strip()
    dob = (data.get('dob') or '').strip()
    house = (data.get('house') or '').strip()
    parents_name = (data.get('parents_name') or '').strip()
    address = (data.get('address') or '').strip()
    status = (data.get('status') or '').strip()

    if not name:
        return jsonify({'status': 'error', 'message': 'Student Name is required.'}), 400
    if not class_:
        return jsonify({'status': 'error', 'message': 'Class is required.'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    existing = c.fetchone()
    if not existing:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student not found.'}), 404

    # Check for duplicate admission number with another student
    if admission_no:
        c.execute("""SELECT student_id FROM students 
                     WHERE LOWER(TRIM(COALESCE(admission_no, '')))=LOWER(TRIM(?))
                       AND student_id != ?""", (admission_no, student_id))
        dup = c.fetchone()
        if dup:
            conn.close()
            return jsonify({'status': 'error', 'message': f'Admission Number "{admission_no}" is already used by another student.'}), 400

    valid_statuses = ('Not Started', 'In Progress', 'Submitted')
    if status and status not in valid_statuses:
        status = existing['status'] or 'Not Started'
    elif not status:
        status = existing['status'] or 'Not Started'

    c.execute("""
        UPDATE students
        SET name=?, admission_no=?, class=?, section=?, dob=?, house=?, parents_name=?, address=?, status=?
        WHERE student_id=?
    """, (name, admission_no, class_, section, dob, house, parents_name, address, status, student_id))

    conn.commit()
    conn.close()

    log_audit_event('admin', 'admin', 'STUDENT_EDIT', 'students', student_id, request.remote_addr)
    return jsonify({
        'status': 'success',
        'message': 'Student details updated successfully.'
    })

@app.route('/api/class_report/subjects')
def api_class_report_subjects():
    """Return list of subjects available across test_papers, questions, and results for a given class & section."""
    class_ = request.args.get('class', '').strip()
    section = request.args.get('section', '').strip()
    if not class_:
        return jsonify({'subjects': []})

    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    conn = get_db()
    c = conn.cursor()

    teacher_id = session.get('teacher_id')
    is_admin = session.get('admin_logged_in')

    if teacher_id and not is_admin:
        c.execute("SELECT subject, is_class_teacher FROM teacher_assignments WHERE teacher_id=? AND (class=? OR class LIKE ?)",
                  (teacher_id, class_, f"%{clean_cls}%"))
        t_rows = [dict(r) for r in c.fetchall()]
        is_ct = any(r.get('is_class_teacher') for r in t_rows)
        t_subjs = sorted(list(set([r['subject'] for r in t_rows if r.get('subject')])))

        if not is_ct and t_subjs:
            conn.close()
            return jsonify({'subjects': t_subjs})

    # Admin or Class Teacher: Query distinct subjects across test_papers, questions, and results
    c.execute("""SELECT DISTINCT subject FROM test_papers WHERE (class=? OR class LIKE ?)
                 UNION
                 SELECT DISTINCT subject FROM questions WHERE (class=? OR class LIKE ?)
                 UNION
                 SELECT DISTINCT subject FROM results WHERE (class=? OR class LIKE ?)""",
              (class_, f"%{clean_cls}%", class_, f"%{clean_cls}%", class_, f"%{clean_cls}%"))
    subjects = sorted(list(set(r['subject'] for r in c.fetchall() if r['subject'])))
    conn.close()
    return jsonify({'subjects': subjects})


@app.route('/api/class_report/test_numbers')
def api_class_report_test_numbers():
    """Return list of test numbers created for a given class, section and subject across test_papers, questions, and results."""
    class_ = request.args.get('class', '').strip()
    section = request.args.get('section', '').strip()
    subject = request.args.get('subject', '').strip()
    if not class_ or not subject:
        return jsonify({'test_numbers': []})

    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    conn = get_db()
    c = conn.cursor()

    # Query created test numbers from test_papers, questions (chapter & test_no), and results
    c.execute("""SELECT DISTINCT test_no as tn FROM test_papers WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no IS NOT NULL AND test_no != ''
                 UNION
                 SELECT DISTINCT chapter as tn FROM questions WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND chapter IS NOT NULL AND chapter != ''
                 UNION
                 SELECT DISTINCT test_no as tn FROM questions WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no IS NOT NULL AND test_no != ''
                 UNION
                 SELECT DISTINCT chapter as tn FROM results WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND chapter IS NOT NULL AND chapter != ''""",
              (class_, f"%{clean_cls}%", subject,
               class_, f"%{clean_cls}%", subject,
               class_, f"%{clean_cls}%", subject,
               class_, f"%{clean_cls}%", subject))
    test_numbers = sorted(list(set(r['tn'] for r in c.fetchall() if r['tn'])))
    conn.close()
    return jsonify({'test_numbers': test_numbers})


@app.route('/admin/student/delete/<student_id>', methods=['DELETE'])
@admin_required
def delete_student(student_id):
    """
    API endpoint to delete a student record.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM students WHERE student_id=?",           (student_id,))
    c.execute("DELETE FROM responses WHERE student_id=?",          (student_id,))
    c.execute("DELETE FROM results WHERE student_id=?",            (student_id,))
    c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
    c.execute("DELETE FROM student_class_lock WHERE student_id=?", (student_id,))  # Bug #3 fix
    c.execute("DELETE FROM reattempt_requests WHERE student_id=?", (student_id,))
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'STUDENT_DELETE', 'students', student_id, request.remote_addr)
    return jsonify({'status': 'deleted', 'message': 'Student deleted successfully.'})


# ══════════════════════════════════════════════════════════════════════════
# BULK STUDENT UPLOAD VIA EXCEL — 3 routes
# ══════════════════════════════════════════════════════════════════════════

@app.route('/admin/students/bulk_upload_page')
@admin_required
def bulk_upload_students_page():
    """Render the bulk student upload review/edit page."""
    return render_template('bulk_upload_students.html')


@app.route('/admin/students/bulk_upload_template')
@admin_required
def bulk_upload_template():
    """
    Stream a downloadable .xlsx Excel template for bulk student upload.
    Headers: Admission No, Name, Class, Section, Date of Birth, House, Parent Name, Address
    """
    import io
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from flask import send_file

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Students"

    # ── column definitions ────────────────────────────────────────────────
    headers = [
        ("Admission No", 18, True,  "Login key — what student types at login (required)"),
        ("Name",         28, True,  "Full name of the student (required)"),
        ("Class",        10, True,  "Plain digit: 1 to 12 (required)"),
        ("Section",      12, True,  "A, B, C or D (required)"),
        ("Date of Birth",16, False, "Optional — e.g. 2010-05-15"),
        ("House",        14, False, "Optional — e.g. Red, Blue"),
        ("Parent Name",  24, False, "Optional — parent/guardian name"),
        ("Address",      32, False, "Optional — full address"),
    ]

    # ── header row styling ────────────────────────────────────────────────
    hdr_fill_req  = PatternFill("solid", fgColor="1A237E")   # dark blue   — required
    hdr_fill_opt  = PatternFill("solid", fgColor="3949AB")   # medium blue — optional
    hdr_font      = Font(bold=True, color="FFFFFF", size=11)
    centre        = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin          = Side(style="thin", color="CCCCCC")
    border        = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, (label, width, required, note) in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font      = hdr_font
        cell.fill      = hdr_fill_req if required else hdr_fill_opt
        cell.alignment = centre
        cell.border    = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 30

    # ── instruction row (row 2) ───────────────────────────────────────────
    inst_notes = [h[3] for h in headers]
    inst_fill  = PatternFill("solid", fgColor="E8EAF6")
    inst_font  = Font(italic=True, color="555588", size=9)
    for col_idx, note in enumerate(inst_notes, start=1):
        cell = ws.cell(row=2, column=col_idx, value=note)
        cell.fill      = inst_fill
        cell.font      = inst_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
    ws.row_dimensions[2].height = 40

    # ── sample data row (row 3) — greyed out ─────────────────────────────
    sample = ["2024001", "Aarav Sharma", "10", "A", "2010-04-15", "Red", "Rajesh Sharma", "12 MG Road, Delhi"]
    sample_font = Font(color="999999", italic=True, size=10)
    sample_fill = PatternFill("solid", fgColor="F5F5F5")
    for col_idx, val in enumerate(sample, start=1):
        cell = ws.cell(row=3, column=col_idx, value=val)
        cell.font   = sample_font
        cell.fill   = sample_fill
        cell.border = border

    # ── freeze top 2 rows ────────────────────────────────────────────────
    ws.freeze_panes = "A3"

    # ── add a note tab ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Instructions")
    instructions = [
        ["Field",        "Required?", "Valid Values / Notes"],
        ["Admission No", "YES",       "Unique per student. Students use this to log in. No spaces."],
        ["Name",         "YES",       "Full name, max ~100 chars."],
        ["Class",        "YES",       "Plain number only: 1, 2, 3 ... 12 (not 'Class 6' or 'VI')."],
        ["Section",      "YES",       "Exactly one letter: A, B, C, or D (uppercase)."],
        ["Date of Birth","No",        "YYYY-MM-DD format e.g. 2010-05-15. Leave blank if unknown."],
        ["House",        "No",        "House name e.g. Red, Blue, Green, Yellow."],
        ["Parent Name",  "No",        "Full name of parent/guardian."],
        ["Address",      "No",        "Full address string."],
    ]
    for r_idx, row_data in enumerate(instructions, start=1):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1A237E")
            cell.border = border
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 55

    # ── stream to client ──────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="bulk_student_upload_template.xlsx"
    )


@app.route('/admin/students/bulk_upload', methods=['POST'])
@admin_required
def bulk_upload_students():
    """
    Parse an uploaded Excel file, validate every row, and return JSON
    describing each row's status — NOTHING is saved to the database here.

    Validation checks per row:
      (a) DB duplicate    — student_id already exists in students table
      (b) Batch duplicate — same student_id appears more than once in the file
      (c) Missing field   — name, class, or section is blank
      (d) Invalid class   — not in '1'..'12'
      (e) Invalid section — not in A/B/C/D
    """
    import openpyxl

    f = request.files.get('excel_file')
    if not f or not f.filename:
        return jsonify({'status': 'error', 'message': 'No file uploaded.'}), 400

    fn = f.filename.lower()
    if not fn.endswith('.xlsx') and not fn.endswith('.xls'):
        return jsonify({'status': 'error', 'message': 'Only .xlsx files are accepted.'}), 400

    try:
        wb = openpyxl.load_workbook(f, data_only=True)
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Cannot read Excel file: {e}'}), 400

    ws = wb.active

    VALID_CLASSES  = {str(i) for i in range(1, 13)}
    VALID_SECTIONS = {'A', 'B', 'C', 'D'}
    EXPECTED_COLS  = ['Admission No', 'Name', 'Class', 'Section',
                      'Date of Birth', 'House', 'Parent Name', 'Address']

    # ── read header row ───────────────────────────────────────────────────
    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter, None)
    if header_row is None:
        return jsonify({'status': 'error', 'message': 'The Excel file is empty.'}), 400

    # Normalise header names (strip whitespace)
    headers = [str(h).strip() if h is not None else '' for h in header_row]

    # Skip instruction row if it looks like a note (second row)
    # We peek at the second row — if first cell doesn't look like an admission_no, skip it
    peek = next(rows_iter, None)
    if peek is not None:
        peek_val = str(peek[0]).strip() if peek[0] is not None else ''
        # If the "Admission No" cell of row 2 starts with lowercase or contains "optional/login/required"
        if peek_val.lower().startswith('login') or 'required' in peek_val.lower() or 'optional' in peek_val.lower():
            peek = None  # discard instruction row, read next

    # ── map column index from headers ─────────────────────────────────────
    def col(name):
        try:
            return headers.index(name)
        except ValueError:
            return -1

    idx_adm  = col('Admission No')
    idx_name = col('Name')
    idx_cls  = col('Class')
    idx_sec  = col('Section')
    idx_dob  = col('Date of Birth')
    idx_hse  = col('House')
    idx_par  = col('Parent Name')
    idx_adr  = col('Address')

    if idx_adm == -1 or idx_name == -1 or idx_cls == -1 or idx_sec == -1:
        return jsonify({
            'status': 'error',
            'message': (
                'Missing required columns. Expected: "Admission No", "Name", "Class", "Section". '
                'Please use the downloaded template.'
            )
        }), 400

    def cell_val(row, idx):
        if idx == -1 or idx >= len(row):
            return ''
        v = row[idx]
        return str(v).strip() if v is not None else ''

    # ── collect all data rows ─────────────────────────────────────────────
    raw_rows = []
    if peek is not None:
        raw_rows.append(peek)
    for row in rows_iter:
        if all((v is None or str(v).strip() == '') for v in row):
            continue  # skip fully blank rows
        raw_rows.append(row)

    if not raw_rows:
        return jsonify({'status': 'error', 'message': 'No data rows found in the file (only header was present).'}), 400

    # ── first pass: extract data + detect batch duplicates ───────────────
    parsed = []
    adm_seen = {}   # admission_no -> first row index (0-based)
    for i, row in enumerate(raw_rows):
        adm     = cell_val(row, idx_adm)
        name    = cell_val(row, idx_name)
        cls     = cell_val(row, idx_cls).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        sec     = cell_val(row, idx_sec).upper().strip()
        dob     = cell_val(row, idx_dob)
        house   = cell_val(row, idx_hse)
        par     = cell_val(row, idx_par)
        addr    = cell_val(row, idx_adr)

        parsed.append({
            'row_index':    i,
            'admission_no': adm,
            'name':         name,
            'class':        cls,
            'section':      sec,
            'dob':          dob,
            'house':        house,
            'parents_name': par,
            'address':      addr,
            'issues':       [],
            'status':       'ok',  # overwritten below
        })

        if adm:
            if adm in adm_seen:
                # mark both this row and the previous occurrence
                adm_seen[adm].append(i)
            else:
                adm_seen[adm] = [i]

    # ── fetch existing student_ids from DB ────────────────────────────────
    conn = get_db()
    c    = conn.cursor()
    batch_adm_list = [r['admission_no'] for r in parsed if r['admission_no']]
    existing_ids   = set()
    if batch_adm_list:
        placeholders = ','.join(['?'] * len(batch_adm_list))
        c.execute(f"SELECT student_id FROM students WHERE student_id IN ({placeholders})", batch_adm_list)
        existing_ids = {row[0] for row in c.fetchall()}
    conn.close()

    # ── second pass: assign issue flags ──────────────────────────────────
    batch_dup_indices = set()
    for adm, indices in adm_seen.items():
        if len(indices) > 1:
            for idx in indices:
                batch_dup_indices.add(idx)

    for r in parsed:
        issues = []
        adm = r['admission_no']

        if not adm:
            issues.append({'field': 'admission_no', 'type': 'missing', 'message': 'Admission No is required'})
        else:
            if adm in existing_ids:
                issues.append({'field': 'admission_no', 'type': 'db_duplicate',
                                'message': 'Already exists in database'})
            if r['row_index'] in batch_dup_indices:
                issues.append({'field': 'admission_no', 'type': 'batch_duplicate',
                                'message': 'Duplicate within this file'})

        if not r['name']:
            issues.append({'field': 'name', 'type': 'missing', 'message': 'Name is required'})
        if not r['class']:
            issues.append({'field': 'class', 'type': 'missing', 'message': 'Class is required'})
        elif r['class'] not in VALID_CLASSES:
            issues.append({'field': 'class', 'type': 'invalid',
                           'message': f'Invalid class "{r["class"]}" — must be 1 to 12'})
        if not r['section']:
            issues.append({'field': 'section', 'type': 'missing', 'message': 'Section is required'})
        elif r['section'] not in VALID_SECTIONS:
            issues.append({'field': 'section', 'type': 'invalid',
                           'message': f'Invalid section "{r["section"]}" — must be A, B, C or D'})

        r['issues'] = issues
        r['status'] = 'ok' if not issues else 'error'

    total_ok    = sum(1 for r in parsed if r['status'] == 'ok')
    total_error = sum(1 for r in parsed if r['status'] == 'error')

    return jsonify({
        'status':      'parsed',
        'rows':        parsed,
        'total':       len(parsed),
        'total_ok':    total_ok,
        'total_error': total_error,
        'existing_ids': list(existing_ids),   # sent to frontend for client-side re-validation
    })


@app.route('/admin/students/bulk_save', methods=['POST'])
@admin_required
def bulk_save_students():
    """
    Save only validated rows into the students table.
    Accepts JSON: { "rows": [ { admission_no, name, class, section, dob, house, parents_name, address }, ... ] }
    Performs a final server-side re-validation before inserting.
    Uses INSERT OR IGNORE so race conditions won't crash the batch.
    Returns counts of saved / skipped rows.
    """
    data = request.get_json(force=True, silent=True) or {}
    rows = data.get('rows', [])

    if not rows:
        return jsonify({'status': 'error', 'message': 'No rows provided.'}), 400

    VALID_CLASSES  = {str(i) for i in range(1, 13)}
    VALID_SECTIONS = {'A', 'B', 'C', 'D'}

    conn = get_db()
    c    = conn.cursor()

    # ── fetch existing IDs for server-side re-validation ─────────────────
    adm_list = [r.get('admission_no', '').strip() for r in rows if r.get('admission_no', '').strip()]
    existing_ids = set()
    if adm_list:
        placeholders = ','.join(['?'] * len(adm_list))
        c.execute(f"SELECT student_id FROM students WHERE student_id IN ({placeholders})", adm_list)
        existing_ids = {row[0] for row in c.fetchall()}

    # ── detect batch duplicates within this save request ─────────────────
    seen_in_batch = set()
    batch_dups    = set()
    for r in rows:
        adm = r.get('admission_no', '').strip()
        if adm:
            if adm in seen_in_batch:
                batch_dups.add(adm)
            else:
                seen_in_batch.add(adm)

    saved   = 0
    skipped = []

    for r in rows:
        adm  = r.get('admission_no', '').strip()
        name = r.get('name', '').strip()
        cls  = r.get('class', '').strip()
        sec  = r.get('section', '').strip().upper()
        dob  = r.get('dob', '').strip()
        house     = r.get('house', '').strip()
        par       = r.get('parents_name', '').strip()
        addr      = r.get('address', '').strip()

        # Final validation gate
        fail_reasons = []
        if not adm:
            fail_reasons.append('missing admission_no')
        if adm in existing_ids:
            fail_reasons.append('already exists in DB')
        if adm in batch_dups:
            fail_reasons.append('duplicate in batch')
        if not name:
            fail_reasons.append('missing name')
        if not cls or cls not in VALID_CLASSES:
            fail_reasons.append(f'invalid class "{cls}"')
        if not sec or sec not in VALID_SECTIONS:
            fail_reasons.append(f'invalid section "{sec}"')

        if fail_reasons:
            skipped.append({'admission_no': adm, 'reasons': fail_reasons})
            continue

        try:
            c.execute("""INSERT OR IGNORE INTO students
                         (student_id, name, class, section, admission_no, dob, house, parents_name, address, status)
                         VALUES (?,?,?,?,?,?,?,?,?,'Not Started')""",
                      (adm, name, cls, sec, adm, dob, house, par, addr))
            if c.rowcount > 0:
                saved += 1
            else:
                skipped.append({'admission_no': adm, 'reasons': ['INSERT OR IGNORE — already inserted by concurrent request']})
        except Exception as e:
            skipped.append({'admission_no': adm, 'reasons': [str(e)]})

    conn.commit()
    conn.close()

    return jsonify({
        'status':  'saved',
        'saved':   saved,
        'skipped': skipped,
    })


@app.route('/admin/monitoring')
@admin_required
def monitoring():
    """
    Render the exam monitoring page showing real-time student status.
    """
    return render_template('monitoring.html')

@app.route('/admin/monitoring/data')
@admin_required
def monitoring_data():
    """
    API endpoint to fetch real-time monitoring data.
    """
    auto_submit_expired_exams()
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT name, class, subject, student_id, ip, status, exam_started_at, current_section FROM students ORDER BY status, name")
    students = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify({'students': students})

@app.route('/admin/evaluate')
@admin_required
def evaluate():
    """
    Render the page for evaluating descriptive responses.
    """
    conn = get_db()
    c = conn.cursor()

    # BUG-003 Fix: fetch negative marking settings FIRST
    c.execute("SELECT negative_marking, negative_value FROM exam_control WHERE id=1")
    ec = c.fetchone()
    neg_enabled = bool(ec['negative_marking']) if ec else False
    neg_value   = float(ec['negative_value'])  if ec else 0.33

    c.execute("SELECT student_id, name, class, section, subject, test_no FROM students WHERE status='Submitted'")
    students = c.fetchall()
    for student in students:
        student_id = student['student_id']
        name  = student['name']
        class_  = student['class']
        section_ = student['section'] or ''
        subject = student['subject']
        test_no = (student['test_no'] or '').strip() if ('test_no' in student.keys() and student['test_no']) else ''
        if test_no:
            c.execute("SELECT id, option_a, option_b, option_c, option_d, correct_answer FROM questions WHERE class=? AND subject=? AND (test_no=? OR chapter=?)", (class_, subject, test_no, test_no))
        else:
            c.execute("SELECT id, option_a, option_b, option_c, option_d, correct_answer FROM questions WHERE class=? AND subject=?", (class_, subject))
        q_rows = c.fetchall()
        correct_map = {}
        for q in q_rows:
            cor_letter = _normalize_correct_answer(
                q['correct_answer'],
                q['option_a'],
                q['option_b'],
                q['option_c'],
                q['option_d']
            )
            if cor_letter:
                correct_map[q['id']] = cor_letter
        c.execute("SELECT question_id, selected_option FROM responses WHERE student_id=?", (student_id,))
        responses = c.fetchall()

        # BUG-003 Fix: apply negative marking to each wrong answer
        raw_score = 0.0
        total     = len(correct_map)
        for resp in responses:
            qid      = resp['question_id']
            selected = (resp['selected_option'] or '').strip().upper()
            if qid not in correct_map:
                continue
            if selected == correct_map[qid]:
                raw_score += 1.0
            elif selected and neg_enabled:
                raw_score -= neg_value

        score      = max(0.0, raw_score)
        percentage = round((score / total * 100), 2) if total > 0 else 0.0
        _save_or_update_result(c, student_id, name, class_, section_, subject, score, total, percentage, test_no)
    conn.commit()
    conn.close()
    return redirect(url_for('results_page'))
@app.route('/api/results/filters')
def api_results_filters():
    """Return distinct classes, sections, subjects, test_numbers from results."""
    conn = get_db()
    c = conn.cursor()

    # Classes from results
    c.execute("SELECT DISTINCT class FROM results WHERE class IS NOT NULL AND class != '' ORDER BY class")
    classes = [r['class'] for r in c.fetchall()]

    # Sections from students (join with results)
    c.execute("""SELECT DISTINCT s.section FROM results r
                 JOIN students s ON r.student_id = s.student_id
                 WHERE s.section IS NOT NULL AND s.section != ''
                 ORDER BY s.section""")
    sections = [r['section'] for r in c.fetchall()]

    # Subjects from results
    c.execute("SELECT DISTINCT subject FROM results WHERE subject IS NOT NULL AND subject != '' ORDER BY subject")
    subjects = [r['subject'] for r in c.fetchall()]

    # Test numbers (chapter) from results
    c.execute("SELECT DISTINCT chapter FROM results WHERE chapter IS NOT NULL AND chapter != '' ORDER BY chapter")
    test_numbers = [r['chapter'] for r in c.fetchall()]

    conn.close()
    return jsonify({
        'classes': classes,
        'sections': sections,
        'subjects': subjects,
        'test_numbers': test_numbers
    })

@app.route('/api/results/summary')
def api_results_summary():
    """Return aggregated summary per class, section, subject, test_no."""
    conn = get_db()
    c = conn.cursor()

    query = """
        SELECT
            r.class,
            COALESCE(r.section, s.section, '') as section,
            r.subject,
            r.chapter as test_no,
            COUNT(DISTINCT r.student_id) as student_count,
            AVG(r.score) as avg_score,
            MAX(r.score) as max_score,
            MIN(r.score) as min_score
        FROM results r
        LEFT JOIN students s ON r.student_id = s.student_id
        WHERE r.class IS NOT NULL AND r.class != ''
          AND r.subject IS NOT NULL AND r.subject != ''
        GROUP BY r.class, COALESCE(r.section, s.section, ''), r.subject, r.chapter
        ORDER BY r.class, COALESCE(r.section, s.section, ''), r.subject, r.chapter
    """
    c.execute(query)
    rows = c.fetchall()
    summary = [dict(row) for row in rows]

    # Total registered students from students table
    c.execute("SELECT COUNT(*) as total FROM students")
    total_st_row = c.fetchone()
    c.execute("SELECT COUNT(DISTINCT student_id) as total_res FROM results")
    total_res_row = c.fetchone()
    total_students = max(total_st_row['total'] if total_st_row else 0, total_res_row['total_res'] if total_res_row else 0)
    conn.close()

    return jsonify({
        'summary': summary,
        'total_students': total_students
    })

@app.route('/api/verify_admin_password', methods=['POST'])
def verify_admin_password():
    """Verify admin password for elevation from results page to admin dashboard."""
    data = request.get_json(silent=True) or request.form or {}
    password = data.get('password', '').strip()
    admin_pass = os.environ.get('ADMIN_PASSWORD', 'admin123').strip()
    if password == admin_pass:
        session['admin_logged_in'] = True
        return jsonify({'status': 'success', 'redirect': url_for('admin_dashboard')})
    return jsonify({'status': 'error', 'message': 'Incorrect admin password.'}), 401

@app.route('/admin/results')
def results_page():
    """
    Render the results page showing all test results. Accessible to both Admin and Teacher roles.
    """
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return redirect(url_for('teacher_login'))
    return render_template('results.html',
                           is_admin=bool(session.get('admin_logged_in')),
                           is_teacher=bool(session.get('teacher_logged_in')))

@app.route('/admin/results/data')
def results_data():
    """
    API endpoint to fetch results data with filtering. Accessible to both Admin and Teacher roles.
    """
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    """
    API endpoint to fetch results data with filtering:
    class, section (via join), subject, test_no (chapter).
    """
    class_filter   = request.args.get('class', '')
    section_filter = request.args.get('section', '')
    subject_filter = request.args.get('subject', '')
    test_filter    = request.args.get('test_no', '')

    conn = get_db()
    c = conn.cursor()

    query = """
        SELECT
            r.id, r.student_id, r.name, r.class, r.subject,
            r.score, r.total_questions, r.percentage, r.test_date,
            r.chapter, r.section_breakdown,
            COALESCE(r.section, s.section, '') as section,
            s.exam_started_at
        FROM results r
        LEFT JOIN students s ON r.student_id = s.student_id
        WHERE r.id IN (
            SELECT MAX(id) FROM results GROUP BY student_id, LOWER(subject), COALESCE(chapter, '')
        )
    """
    params = []

    if class_filter:
        query += " AND r.class = ?"
        params.append(class_filter)
    if section_filter:
        query += " AND COALESCE(r.section, s.section, '') = ?"
        params.append(section_filter)
    if subject_filter:
        query += " AND r.subject = ?"
        params.append(subject_filter)
    if test_filter:
        query += " AND r.chapter = ?"
        params.append(test_filter)

    query += " ORDER BY r.test_date DESC, r.class, r.subject, r.name"
    c.execute(query, params)
    results = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(results)
@app.route('/admin/results_filter')
@admin_required
def results_filter():
    """
    Render the advanced results filter and export page.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class FROM results WHERE class IS NOT NULL AND class != '' UNION SELECT DISTINCT class FROM questions WHERE class IS NOT NULL AND class != '' ORDER BY class")
    classes = [dict(r)['class'] for r in c.fetchall()]
    
    c.execute("SELECT DISTINCT subject FROM results WHERE subject IS NOT NULL AND subject != '' UNION SELECT DISTINCT subject FROM questions WHERE subject IS NOT NULL AND subject != '' ORDER BY subject")
    subjects = [dict(r)['subject'] for r in c.fetchall()]
    
    c.execute("SELECT DISTINCT section FROM students WHERE section IS NOT NULL AND section != '' ORDER BY section")
    sections = [dict(r)['section'] for r in c.fetchall()]
    conn.close()
    return render_template('results_filter.html', classes=classes, subjects=subjects, sections=sections)

@app.route('/admin/export_results_page')
@admin_required
def export_results_page():
    """
    Render the page for exporting results.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class FROM students ORDER BY class")
    classes = [row['class'] for row in c.fetchall()]
    c.execute("SELECT DISTINCT subject FROM students ORDER BY subject")
    subjects = [row['subject'] for row in c.fetchall()]
    c.execute("SELECT DISTINCT section FROM students WHERE section IS NOT NULL AND section != '' ORDER BY section")
    sections = [row['section'] for row in c.fetchall()]
    conn.close()
    return render_template('export_results_filter.html', classes=classes, subjects=subjects, sections=sections)

@app.route('/admin/export/results')
@admin_required
def export_results():
    """
    Export results to Excel, PDF, or JSON (preview) with filters and sorting.
    """
    # ── Parse filters & sorting ────────────────────────────────────────
    class_filter   = request.args.get('class', '')
    section_filter = request.args.get('section', '')
    subject_filter = request.args.get('subject', '')
    test_filter    = request.args.get('test_no', '')
    start_date     = request.args.get('start_date', '')
    end_date       = request.args.get('end_date', '')
    single_date    = request.args.get('date', '')
    format_type    = request.args.get('format', 'excel')  # 'excel', 'pdf', 'json'
    sort_by        = request.args.get('sort_by', 'rank')   # 'rank', 'name', 'student_id', 'score', 'percentage', 'class', 'test_date'
    sort_order     = request.args.get('sort_order', 'asc') # 'asc', 'desc'

    conn = get_db()
    c = conn.cursor()

    # ── Build query ────────────────────────────────────────────────────
    query = '''
        SELECT
            s.student_id,
            s.name,
            s.class,
            COALESCE(s.section, '') as section,
            s.subject,
            COALESCE(NULLIF(r.chapter,''), NULLIF(s.test_no,''), '—') as test_no,
            COALESCE(r.score, 0) as score,
            COALESCE(r.total_questions, 0) as total_questions,
            COALESCE(r.percentage, 0.0) as percentage,
            COALESCE(r.test_date, s.exam_started_at) as test_date,
            s.exam_started_at
        FROM students s
        LEFT JOIN results r ON r.id = (
            SELECT id FROM results 
            WHERE student_id = s.student_id 
              AND (LOWER(subject) = LOWER(s.subject) OR subject IS NULL OR subject = '')
              AND (chapter = s.test_no OR chapter IS NULL OR chapter = '' OR s.test_no IS NULL OR s.test_no = '')
            ORDER BY id DESC LIMIT 1
        )
        WHERE 1=1
    '''
    params = []

    if class_filter and class_filter != 'All':
        clean_cls = class_filter.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        query += " AND (s.class = ? OR s.class LIKE ? OR r.class = ? OR r.class LIKE ?)"
        params.extend([class_filter, f"%{clean_cls}%", class_filter, f"%{clean_cls}%"])

    if section_filter and section_filter != 'All':
        query += " AND LOWER(COALESCE(s.section, '')) = LOWER(?)"
        params.append(section_filter.strip())

    if subject_filter and subject_filter != 'All':
        query += " AND LOWER(COALESCE(s.subject, r.subject)) = LOWER(?)"
        params.append(subject_filter.strip())

    if test_filter and test_filter != '' and test_filter != 'All':
        query += " AND (r.chapter = ? OR s.test_no = ?)"
        params.extend([test_filter, test_filter])

    if single_date:
        query += " AND DATE(COALESCE(r.test_date, s.exam_started_at)) = ?"
        params.append(single_date)
    else:
        if start_date:
            query += " AND DATE(COALESCE(r.test_date, s.exam_started_at)) >= ?"
            params.append(start_date)
        if end_date:
            query += " AND DATE(COALESCE(r.test_date, s.exam_started_at)) <= ?"
            params.append(end_date)

    query += " ORDER BY LOWER(TRIM(s.name)) ASC, s.student_id ASC"

    c.execute(query, params)
    rows = c.fetchall()
    
    seen = set()
    deduped_results = []
    for row in rows:
        r_dict = dict(row)
        key = (r_dict['student_id'], r_dict['subject'], r_dict['test_no'])
        if key not in seen:
            seen.add(key)
            deduped_results.append(r_dict)
            
    results = deduped_results
    conn.close()

    # ─── Compute Ranks (based on score & percentage DESC) ───────────────
    rank_pool = sorted(
        results,
        key=lambda x: (float(x.get('percentage') or 0.0), float(x.get('score') or 0.0)),
        reverse=True
    )
    current_rank = 0
    prev_metric = None
    for idx, item in enumerate(rank_pool):
        metric = (round(float(item.get('percentage') or 0.0), 2), round(float(item.get('score') or 0.0), 2))
        if metric != prev_metric:
            current_rank = idx + 1
            prev_metric = metric
        item['rank'] = current_rank

    # ─── Multi-column Sorting ───────────────────────────────────────────
    is_desc = (sort_order.lower() == 'desc')

    if sort_by == 'rank':
        results.sort(key=lambda x: (int(x.get('rank') or 999999), (x.get('name') or '').strip().lower()), reverse=is_desc)
    elif sort_by == 'student_id':
        results.sort(key=lambda x: (x.get('student_id') or '').strip().lower(), reverse=is_desc)
    elif sort_by == 'name':
        results.sort(key=lambda x: (x.get('name') or '').strip().lower(), reverse=is_desc)
    elif sort_by == 'percentage':
        results.sort(key=lambda x: (float(x.get('percentage') or 0.0), float(x.get('score') or 0.0)), reverse=is_desc)
    elif sort_by == 'score':
        results.sort(key=lambda x: (float(x.get('score') or 0.0), float(x.get('percentage') or 0.0)), reverse=is_desc)
    elif sort_by == 'class':
        results.sort(key=lambda x: (str(x.get('class') or '').lower(), str(x.get('section') or '').lower(), (x.get('name') or '').lower()), reverse=is_desc)
    elif sort_by == 'subject':
        results.sort(key=lambda x: ((x.get('subject') or '').lower(), (x.get('name') or '').lower()), reverse=is_desc)
    elif sort_by == 'test_no':
        results.sort(key=lambda x: ((x.get('test_no') or '').lower(), int(x.get('rank') or 999999)), reverse=is_desc)
    elif sort_by == 'test_date':
        results.sort(key=lambda x: str(x.get('test_date') or x.get('exam_started_at') or ''), reverse=is_desc)
    else:
        results.sort(key=lambda x: (int(x.get('rank') or 999999), (x.get('name') or '').strip().lower()), reverse=is_desc)

    total_students = len(results)

    # ─── JSON (Preview) ──────────────────────────────────────────────────
    if format_type == 'json':
        return jsonify({'results': results, 'count': total_students})

    # ─── Common metadata ────────────────────────────────────────────────
    school_name = get_setting('school_name', 'RRB Group of Schools')
    class_display = class_filter if class_filter and class_filter != 'All' else "All Classes"
    section_display = section_filter if section_filter and section_filter != 'All' else "All Sections"
    subject_display = subject_filter if subject_filter and subject_filter != 'All' else "All Subjects"
    test_display = test_filter if test_filter else "All Tests"

    if single_date:
        date_display = single_date
    elif start_date and end_date:
        date_display = f"{start_date} to {end_date}"
    elif start_date:
        date_display = f"From {start_date}"
    elif end_date:
        date_display = f"Up to {end_date}"
    else:
        date_display = "All Dates"

    export_time = datetime.datetime.now().strftime('%d-%m-%Y %H:%M')

    # ─── PDF Export ──────────────────────────────────────────────────────
    if format_type == 'pdf':
        # Build logo base64 if exists
        logo_path = get_setting('logo_path', '')
        logo_base64 = None
        if logo_path:
            abs_path = os.path.join(app.static_folder, logo_path)
            if os.path.exists(abs_path):
                with open(abs_path, "rb") as f:
                    logo_bytes = f.read()
                    ext = os.path.splitext(abs_path)[1].lower()
                    mime = "image/png" if ext == ".png" else "image/jpeg"
                    logo_base64 = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode()}"

        rendered = render_template('export_results_pdf.html',
            school_name=school_name,
            logo_base64=logo_base64,
            class_display=class_display,
            section_display=section_display,
            subject_display=subject_display,
            test_display=test_display,
            date_display=date_display,
            export_time=export_time,
            total_students=total_students,
            results=results
        )
        if request.args.get('download') == 'pdf':
            pdf = HTML(string=rendered).write_pdf()
            response = make_response(pdf)
            response.headers['Content-Type'] = 'application/pdf'
            response.headers['Content-Disposition'] = f'attachment; filename=export_results_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
            return response

        return rendered

    # ─── Excel Export (default) ─────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exam Results"

    # Header rows
    ws.merge_cells('A1:I1')
    ws['A1'].value = school_name
    ws['A1'].font = openpyxl.styles.Font(bold=True, size=16)
    ws['A1'].alignment = openpyxl.styles.Alignment(horizontal='center')

    ws.merge_cells('A2:I2')
    ws['A2'].value = f"Class: {class_display}  |  Section: {section_display}  |  Subject: {subject_display}  |  Test: {test_display}"
    ws['A2'].font = openpyxl.styles.Font(bold=True, size=12)
    ws['A2'].alignment = openpyxl.styles.Alignment(horizontal='center')

    ws.merge_cells('A3:I3')
    ws['A3'].value = f"Date: {date_display}  |  Total Students: {total_students}  |  Sorted By: {sort_by.upper()} ({sort_order.upper()})  |  Generated: {export_time}"
    ws['A3'].font = openpyxl.styles.Font(italic=True)
    ws['A3'].alignment = openpyxl.styles.Alignment(horizontal='center')
    ws.append([])

    headers = ['Rank', 'Student ID', 'Name', 'Class', 'Section', 'Subject', 'Test No', 'Score', 'Percentage']
    header_row = 5
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
        cell.alignment = openpyxl.styles.Alignment(horizontal='center')

    for row_idx, row in enumerate(results, start=header_row + 1):
        ws.cell(row=row_idx, column=1, value=row.get('rank', '—'))
        ws.cell(row=row_idx, column=2, value=row['student_id'])
        ws.cell(row=row_idx, column=3, value=row['name'])
        ws.cell(row=row_idx, column=4, value=row['class'])
        ws.cell(row=row_idx, column=5, value=row.get('section', 'N/A'))
        ws.cell(row=row_idx, column=6, value=row['subject'])
        ws.cell(row=row_idx, column=7, value=row.get('test_no', 'N/A'))
        ws.cell(row=row_idx, column=8, value=f"{row['score']}/{row['total_questions']}" if row['total_questions'] else row['score'])
        ws.cell(row=row_idx, column=9, value=f"{row['percentage']:.1f}%" if row['percentage'] else 'N/A')

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 12

    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    file_parts = []
    file_parts.append(class_filter if class_filter and class_filter != 'All' else 'all')
    file_parts.append(subject_filter if subject_filter and subject_filter != 'All' else 'all')
    if start_date and end_date:
        file_parts.append(f"{start_date}_to_{end_date}")
    elif start_date:
        file_parts.append(f"from_{start_date}")
    elif end_date:
        file_parts.append(f"to_{end_date}")
    else:
        file_parts.append("all_dates")
    file_parts.append(f"sort_{sort_by}_{sort_order}")
    file_parts.append(timestamp)
    filename = f"results_{'_'.join(file_parts)}.xlsx"

    filepath = os.path.join(app.config['EXPORT_FOLDER'], filename)
    wb.save(filepath)

    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/admin/student/responses/<student_id>')
@admin_required
def view_student_responses(student_id):
    """
    Render page to view all responses of a specific student.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        return "Student not found", 404
    log_audit_event('admin', 'admin', 'STUDENT_RESULT_VIEW', 'students', student_id, request.remote_addr)

    # Fetch questions, prioritizing shuffled_questions for exact student exam sequence & multi-section support
    c.execute("""
        SELECT q.id, q.question, q.option_a, q.option_b, q.option_c, q.option_d, q.correct_answer,
               q.section_id, ts.section_name
        FROM shuffled_questions sq
        JOIN questions q ON sq.question_id = q.id
        LEFT JOIN test_sections ts ON q.section_id = ts.id
        WHERE sq.student_id = ?
        ORDER BY sq.shuffled_index
    """, (student_id,))
    questions = c.fetchall()

    test_no = (student['test_no'] if ('test_no' in student.keys() and student['test_no']) else '').strip()
    if not questions:
        if test_no:
            c.execute("SELECT id, question, option_a, option_b, option_c, option_d, correct_answer, NULL as section_id, '' as section_name FROM questions WHERE class=? AND subject=? AND (test_no=? OR chapter=?) ORDER BY id",
                      (student['class'], student['subject'], test_no, test_no))
        else:
            c.execute("SELECT id, question, option_a, option_b, option_c, option_d, correct_answer, NULL as section_id, '' as section_name FROM questions WHERE class=? AND subject=? ORDER BY id",
                      (student['class'], student['subject']))
        questions = c.fetchall()

    c.execute("SELECT section_breakdown FROM results WHERE student_id=? ORDER BY id DESC LIMIT 1", (student_id,))
    res_row = c.fetchone()
    section_breakdown = []
    if res_row and res_row['section_breakdown']:
        try:
            section_breakdown = json.loads(res_row['section_breakdown'])
        except Exception:
            section_breakdown = []

    c.execute("SELECT question_id, selected_option FROM responses WHERE student_id=?", (student_id,))
    responses = {row['question_id']: row['selected_option'] for row in c.fetchall()}

    # Fetch coding submissions & approved coding questions for this student
    clean_cls = (student['class'] or '').replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    c.execute("SELECT * FROM coding_submissions WHERE student_id=? ORDER BY question_id", (student_id,))
    coding_subs = {row['question_id']: dict(row) for row in c.fetchall()}

    c.execute("""SELECT id, title, problem_statement, constraints, sample_input, sample_output, total_marks 
                 FROM coding_questions 
                 WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND (test_no=? OR topic=? OR test_no='') AND approved=1
                 ORDER BY id""",
              (student['class'], f"%{clean_cls}%", student['subject'], test_no, test_no))
    coding_qs = [dict(r) for r in c.fetchall()]

    conn.close()

    question_data = []
    for q in questions:
        selected = responses.get(q['id'], '')
        correct_letter = _normalize_correct_answer(
            q['correct_answer'],
            q['option_a'],
            q['option_b'],
            q['option_c'],
            q['option_d']
        )
        is_correct = bool(selected and correct_letter and selected.strip().upper() == correct_letter)

        sec_name = ''
        try:
            sec_name = q['section_name'] or ''
        except Exception:
            pass

        question_data.append({
            'id': q['id'],
            'question': q['question'],
            'section_name': sec_name,
            'options': {
                'A': q['option_a'],
                'B': q['option_b'],
                'C': q['option_c'],
                'D': q['option_d']
            },
            'correct': correct_letter or '?',  # display the letter, not raw text
            'selected': selected,
            'is_correct': is_correct
        })

    coding_responses = []
    for q in coding_qs:
        sub = coding_subs.get(q['id'])
        case_results = []
        if sub and sub.get('stdout_output'):
            try:
                case_results = json.loads(sub['stdout_output'])
            except Exception:
                case_results = []
                
        coding_responses.append({
            'id': q['id'],
            'title': q['title'],
            'problem_statement': q['problem_statement'],
            'total_marks': float(q['total_marks'] or 10.0),
            'submitted': bool(sub),
            'language': sub['language'] if sub else 'N/A',
            'code': sub['code'] if sub else '',
            'status': sub['status'] if sub else 'Unattempted',
            'test_cases_passed': sub['test_cases_passed'] if sub else 0,
            'total_test_cases': sub['total_test_cases'] if sub else 0,
            'marks_awarded': float(sub['marks_awarded']) if sub else 0.0,
            'case_results': case_results
        })

    return render_template('student_responses.html', student=student, questions=question_data, coding_responses=coding_responses, section_breakdown=section_breakdown)

# ─────── HIERARCHICAL RESULTS FILTERING APIS ────────

@app.route('/api/admin/results/get_classes')
@admin_required
def api_get_classes():
    """
    API endpoint to get all classes with results.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class FROM results WHERE class IS NOT NULL AND class != '' ORDER BY class")
    classes = [row['class'] for row in c.fetchall()]
    conn.close()
    return jsonify({'classes': classes})

@app.route('/api/admin/results/get_subjects/<class_name>')
@admin_required
def api_get_subjects(class_name):
    """
    API endpoint to get subjects for a specific class.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT subject FROM results WHERE class=? AND subject IS NOT NULL AND subject != '' ORDER BY subject", (class_name,))
    subjects = [row['subject'] for row in c.fetchall()]
    conn.close()
    return jsonify({'subjects': subjects})

@app.route('/api/admin/results/get_test_numbers/<class_name>/<subject_name>')
@admin_required
def api_get_test_numbers(class_name, subject_name):
    """
    API endpoint to get test numbers for a specific class and subject.
    """
    conn = get_db()
    c = conn.cursor()
    # Get test_no from chapters (stored in results.chapter)
    c.execute("SELECT DISTINCT chapter FROM results WHERE class=? AND subject=? AND chapter IS NOT NULL AND chapter != '' ORDER BY chapter", 
              (class_name, subject_name))
    test_numbers = [row['chapter'] for row in c.fetchall()]
    conn.close()
    return jsonify({'test_numbers': test_numbers})

@app.route('/api/admin/results/get_dates/<class_name>/<subject_name>/<test_no>')
@admin_required
def api_get_dates(class_name, subject_name, test_no):
    """
    API endpoint to get test dates for a specific class, subject, and test number.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT DISTINCT DATE(test_date) as date FROM results 
                 WHERE class=? AND subject=? AND chapter=? 
                 ORDER BY test_date DESC""", 
              (class_name, subject_name, test_no))
    dates = [row['date'] for row in c.fetchall()]
    conn.close()
    return jsonify({'dates': dates})

@app.route('/api/admin/results/get_filtered_results')
@admin_required
def api_get_filtered_results():
    """
    API endpoint to get filtered results based on class, subject, test_no, and date.
    """
    class_name = request.args.get('class', '')
    subject_name = request.args.get('subject', '')
    test_no = request.args.get('test_no', '')
    date_filter = request.args.get('date', '')

    conn = get_db()
    c = conn.cursor()

    query = '''SELECT r.id, r.student_id, r.name, r.class, r.subject, r.score,
                      r.total_questions, r.percentage, r.test_date, r.chapter
               FROM results r
               WHERE 1=1'''
    params = []

    if class_name:
        query += " AND r.class = ?"
        params.append(class_name)
    if subject_name:
        query += " AND r.subject = ?"
        params.append(subject_name)
    if test_no:
        query += " AND r.chapter = ?"
        params.append(test_no)
    if date_filter:
        query += " AND DATE(r.test_date) = ?"
        params.append(date_filter)

    query += " ORDER BY r.name ASC"

    c.execute(query, params)
    results = [dict(row) for row in c.fetchall()]
    conn.close()

    return jsonify({
        'results': results,
        'count': len(results),
        'filters': {
            'class': class_name,
            'subject': subject_name,
            'test_no': test_no,
            'date': date_filter
        }
    })

@app.route('/admin/export/answers')
@admin_required
def export_answers():
    conn = get_db()
    c = conn.cursor()
    c.execute('''SELECT r.student_id, q.question, r.selected_option, q.correct_answer
                 FROM responses r JOIN questions q ON r.question_id = q.id
                 ORDER BY r.student_id, q.id''')
    rows = c.fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Answers"
    ws.append(['Student ID', 'Question', 'Selected', 'Correct'])
    for row in rows:
        ws.append([row['student_id'], row['question'], row['selected_option'], row['correct_answer']])
    filename = f"answers_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = os.path.join(app.config['EXPORT_FOLDER'], filename)
    wb.save(filepath)
    return send_file(filepath, as_attachment=True)

@app.route('/admin/current_test_sessions')
@admin_required
def admin_current_test_sessions():
    """Page showing list of recent test sessions with generate buttons."""
    return render_template('current_test_sessions.html')

@app.route('/api/admin/recent_test_sessions')
@admin_required
def api_recent_test_sessions():
    """Return enriched list of completed and live test sessions with KPIs and filters."""
    conn = get_db()
    c = conn.cursor()
    
    # ── 1. Completed test sessions with full metrics ───────────────────
    c.execute("""
        SELECT 
            r.class,
            COALESCE(NULLIF(r.section, ''), NULLIF(s.section, ''), '') as section,
            r.subject,
            COALESCE(NULLIF(r.chapter, ''), 'Test') as test_no,
            DATE(r.test_date) as session_date,
            MAX(r.test_date) as test_date,
            COUNT(DISTINCT r.student_id) as student_count,
            ROUND(AVG(r.score), 1) as avg_score,
            ROUND(AVG(r.percentage), 1) as avg_percentage,
            MAX(r.score) as max_score,
            MIN(r.score) as min_score,
            MAX(r.total_questions) as total_questions,
            SUM(CASE WHEN r.percentage >= 40 THEN 1 ELSE 0 END) as pass_count
        FROM results r
        LEFT JOIN students s ON r.student_id = s.student_id
        WHERE r.class IS NOT NULL AND r.class != ''
        GROUP BY r.class, COALESCE(NULLIF(r.section, ''), NULLIF(s.section, ''), ''), r.subject, r.chapter, DATE(r.test_date)
        ORDER BY MAX(r.test_date) DESC
        LIMIT 100
    """)
    completed_sessions = []
    for row in c.fetchall():
        s_dict = dict(row)
        total_students = s_dict.get('student_count') or 0
        pass_count = s_dict.get('pass_count') or 0
        s_dict['pass_rate'] = round((pass_count / total_students * 100), 1) if total_students > 0 else 0.0
        completed_sessions.append(s_dict)

    # ── 2. Live active sessions (currently ongoing) ─────────────────────
    c.execute("""
        SELECT
            class,
            COALESCE(section, '') as section,
            subject,
            COALESCE(test_no, 'Test') as test_no,
            DATE(exam_started_at) as session_date,
            MIN(exam_started_at) as start_time,
            COUNT(*) as total_students,
            SUM(CASE WHEN status IN ('In Progress', 'Logged In', 'Started') THEN 1 ELSE 0 END) as active_count,
            SUM(CASE WHEN status = 'Submitted' THEN 1 ELSE 0 END) as submitted_count,
            SUM(CASE WHEN status = 'Not Started' OR status IS NULL THEN 1 ELSE 0 END) as pending_count
        FROM students
        WHERE exam_started_at IS NOT NULL AND status IN ('In Progress', 'Logged In', 'Started')
        GROUP BY class, COALESCE(section, ''), subject, COALESCE(test_no, 'Test')
        ORDER BY MIN(exam_started_at) DESC
    """)
    live_sessions = [dict(row) for row in c.fetchall()]

    # ── 3. Filters available in DB ─────────────────────────────────────
    c.execute("SELECT DISTINCT class FROM results WHERE class IS NOT NULL AND class != '' ORDER BY class")
    classes = [r['class'] for r in c.fetchall()]

    c.execute("""SELECT DISTINCT COALESCE(NULLIF(r.section, ''), NULLIF(s.section, ''), '') as sec 
                 FROM results r LEFT JOIN students s ON r.student_id = s.student_id 
                 WHERE COALESCE(NULLIF(r.section, ''), NULLIF(s.section, ''), '') != '' ORDER BY sec""")
    sections = [r['sec'] for r in c.fetchall()]

    c.execute("SELECT DISTINCT subject FROM results WHERE subject IS NOT NULL AND subject != '' ORDER BY subject")
    subjects = [r['subject'] for r in c.fetchall()]

    # ── 4. Overall KPIs ────────────────────────────────────────────────
    total_students_tested = sum(s.get('student_count', 0) for s in completed_sessions)
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    today_sessions = sum(1 for s in completed_sessions if s.get('session_date') == today_str)
    active_live_students = sum(s.get('active_count', 0) for s in live_sessions)
    avg_pass_rate = round(sum(s.get('pass_rate', 0.0) for s in completed_sessions) / len(completed_sessions), 1) if completed_sessions else 0.0

    conn.close()

    return jsonify({
        'sessions': completed_sessions,
        'live_sessions': live_sessions,
        'filters': {
            'classes': classes,
            'sections': sections,
            'subjects': subjects
        },
        'kpis': {
            'total_sessions': len(completed_sessions),
            'today_sessions': today_sessions,
            'total_students_tested': total_students_tested,
            'active_live_students': active_live_students,
            'avg_pass_rate': avg_pass_rate
        }
    })

@app.route('/api/admin/session_students')
@admin_required
def api_session_students():
    """API endpoint to get list of students and detailed scores for a single test session."""
    class_filter = request.args.get('class', '').strip()
    section_filter = request.args.get('section', '').strip()
    subject_filter = request.args.get('subject', '').strip()
    test_no_filter = request.args.get('test_no', '').strip()
    date_filter = request.args.get('date', '').strip()
    
    conn = get_db()
    c = conn.cursor()
    clean_cls = class_filter.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    
    query = """
        SELECT r.*, COALESCE(NULLIF(s.admission_no, ''), s.student_id) as admission_no,
               COALESCE(NULLIF(r.section, ''), s.section, '') as student_section
        FROM results r
        LEFT JOIN students s ON r.student_id = s.student_id
        WHERE 1=1
    """
    params = []
    if class_filter:
        query += " AND (r.class = ? OR r.class = ? OR r.class LIKE ?)"
        params.extend([class_filter, clean_cls, f"%{clean_cls}%"])
    if section_filter:
        query += " AND (r.section = ? OR s.section = ?)"
        params.extend([section_filter, section_filter])
    if subject_filter:
        query += " AND (LOWER(r.subject) = LOWER(?) OR r.subject LIKE ?)"
        params.extend([subject_filter, f"%{subject_filter}%"])
    if test_no_filter:
        query += " AND (r.chapter = ? OR r.chapter LIKE ?)"
        params.extend([test_no_filter, f"%{test_no_filter}%"])
    if date_filter:
        query += " AND DATE(r.test_date) = DATE(?)"
        params.append(date_filter)
        
    query += " ORDER BY r.score DESC, r.percentage DESC"
    c.execute(query, params)
    results = [dict(row) for row in c.fetchall()]
    conn.close()
    
    for i, res in enumerate(results):
        res['rank'] = i + 1
        
    return jsonify({'students': results, 'count': len(results)})

@app.route('/admin/current_test_results')
@admin_required
def current_test_results():
    """Show results of the test session, filtered by class, section, subject, test_no, date."""
    format_type = request.args.get('format', 'html')
    class_filter = request.args.get('class', '').strip()
    section_filter = request.args.get('section', '').strip()
    subject_filter = request.args.get('subject', '').strip()
    test_no_filter = request.args.get('test_no', '').strip()
    date_filter = request.args.get('date', '').strip()
    
    conn = get_db()
    c = conn.cursor()

    clean_cls = class_filter.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()

    # Build base query with joins
    query = """
        SELECT r.*, COALESCE(NULLIF(s.admission_no, ''), s.student_id) as admission_no,
               COALESCE(NULLIF(r.section, ''), s.section, '') as student_section
        FROM results r
        LEFT JOIN students s ON r.student_id = s.student_id
        WHERE 1=1
    """
    params = []

    if class_filter:
        query += " AND (r.class = ? OR r.class = ? OR r.class LIKE ?)"
        params.extend([class_filter, clean_cls, f"%{clean_cls}%"])
    if section_filter:
        query += " AND (r.section = ? OR s.section = ?)"
        params.extend([section_filter, section_filter])
    if subject_filter:
        query += " AND (LOWER(r.subject) = LOWER(?) OR r.subject LIKE ?)"
        params.extend([subject_filter, f"%{subject_filter}%"])
    if test_no_filter:
        query += " AND (r.chapter = ? OR r.chapter LIKE ?)"
        params.extend([test_no_filter, f"%{test_no_filter}%"])
    if date_filter:
        query += " AND DATE(r.test_date) = DATE(?)"
        params.append(date_filter)

    query += " ORDER BY r.score DESC, r.percentage DESC"
    c.execute(query, params)
    results = [dict(row) for row in c.fetchall()]
    conn.close()

    if not results:
        return "No results found for the selected filters.", 404

    # Assign ranks and parse section_breakdown
    for i, res in enumerate(results):
        res['rank'] = i + 1
        bd = res.get('section_breakdown')
        if bd and isinstance(bd, str):
            try:
                res['parsed_breakdown'] = json.loads(bd)
            except Exception:
                res['parsed_breakdown'] = []
        elif isinstance(bd, list):
            res['parsed_breakdown'] = bd
        else:
            res['parsed_breakdown'] = []
        if res.get('parsed_breakdown'):
            res_max = sum(s.get('max_marks', 0) for s in res['parsed_breakdown'])
            if res_max > 0:
                res['max_marks'] = res_max

    total_q = results[0].get('total_questions', 0) if results else 0
    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')
    logo_base64 = None
    if logo_path:
        abs_path = os.path.join(app.static_folder, logo_path)
        if os.path.exists(abs_path):
            with open(abs_path, "rb") as f:
                logo_bytes = f.read()
                ext = os.path.splitext(abs_path)[1].lower()
                mime = "image/png" if ext == ".png" else "image/jpeg"
                logo_base64 = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode()}"

    total_q_map = {f"{res['class']}_{res['subject']}": res.get('total_questions', 0) for res in results}
    max_total_marks = results[0].get('max_marks') if (results and results[0].get('max_marks')) else total_q

    display_class = results[0]['class'] if results else '—'
    display_section = section_filter or results[0].get('student_section', '—') if results else '—'
    display_subject = results[0]['subject'] if results else '—'
    display_test_no = results[0].get('chapter', '—') if results else '—'
    start_time = results[0]['test_date'] if results else ''
    date_val = date_filter or (start_time[:10] if start_time else '')

    # Compute additional statistics
    scores = [r.get('score', 0) for r in results]
    highest_score = max(scores) if scores else 0
    lowest_score = min(scores) if scores else 0
    avg_score = round(sum(scores) / len(scores), 1) if scores else 0.0
    pass_students = sum(1 for r in results if (r.get('percentage') or 0) >= 40)
    pass_rate = round((pass_students / len(results)) * 100, 1) if results else 0.0

    if format_type == 'pdf':
        rendered = render_template('current_test_results.html',
                                   results=results,
                                   school_name=school_name,
                                   logo_base64=logo_base64,
                                   start_time=start_time,
                                   total_students=len(results),
                                   total_q_map=total_q_map,
                                   max_total_marks=max_total_marks,
                                   display_class=display_class,
                                   display_section=display_section,
                                   display_subject=display_subject,
                                   display_test_no=display_test_no,
                                   date_filter=date_val,
                                   highest_score=highest_score,
                                   lowest_score=lowest_score,
                                   avg_score=avg_score,
                                   pass_rate=pass_rate,
                                   datetime=datetime)
        pdf = HTML(string=rendered).write_pdf()
        response = make_response(pdf)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=current_test_results_{display_class}_{display_subject}.pdf'
        return response
    
    elif format_type == 'excel':
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Results"
        ws.append(['Rank', 'S.No.', 'Admission No', 'Student Name', 'Class', 'Section', 'Subject', 'Test No', f'Marks (out of {max_total_marks})', 'Percentage'])
        for i, res in enumerate(results):
            pct_str = f"{res['percentage']:.1f}%" if res.get('percentage') else 'N/A'
            ws.append([res['rank'], i+1, res.get('admission_no', ''), res['name'], res['class'], res.get('student_section', '—'), res['subject'], res.get('chapter') or display_test_no or '—', f"{res['score']}/{max_total_marks}", pct_str])
        filename = f"current_test_results_{display_class}_{display_subject}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(app.config['EXPORT_FOLDER'], filename)
        wb.save(filepath)
        return send_file(filepath, as_attachment=True, download_name=filename)
    
    else:
        return render_template('current_test_results.html',
                               results=results,
                               school_name=school_name,
                               logo_base64=logo_base64,
                               start_time=start_time,
                               total_students=len(results),
                               total_q_map=total_q_map,
                               max_total_marks=max_total_marks,
                               display_class=display_class,
                               display_section=display_section,
                               display_subject=display_subject,
                               display_test_no=display_test_no,
                               date_filter=date_val,
                               highest_score=highest_score,
                               lowest_score=lowest_score,
                               avg_score=avg_score,
                               pass_rate=pass_rate,
                               datetime=datetime)

       
@app.route('/admin/reattempt_requests')
@admin_required
def admin_reattempt_requests():
    return render_template('reattempt_requests.html', role='admin', api_base='/admin')

@app.route('/admin/reattempt_requests/data')
@admin_required
def reattempt_requests_data():
    filter_status = request.args.get('status', '')
    conn = get_db()
    c = conn.cursor()
    query = """
        SELECT r.id, r.student_id, s.name, r.class, r.subject, r.status, r.requested_at, r.reviewed_at, r.admin_note
        FROM reattempt_requests r
        JOIN students s ON r.student_id = s.student_id
        WHERE 1=1
    """
    params = []
    if filter_status:
        query += " AND r.status = ?"
        params.append(filter_status)
    query += " ORDER BY r.requested_at DESC"
    c.execute(query, params)
    requests = [dict(row) for row in c.fetchall()]
    c.execute("SELECT is_active FROM exam_control WHERE id=1")
    exam_row = c.fetchone()
    exam_is_active = bool(exam_row['is_active']) if exam_row else False
    conn.close()
    return jsonify({'requests': requests, 'exam_is_active': exam_is_active})

@app.route('/teacher/reattempt_requests')
@teacher_required
def teacher_reattempt_requests():
    """
    Render the teacher reattempt requests management page.
    """
    return render_template('reattempt_requests.html', role='teacher', api_base='/teacher')

@app.route('/teacher/reattempt_requests/data')
@teacher_required
def teacher_reattempt_requests_data():
    """
    Fetch reattempt requests for students in the teacher's allotted classes and subjects.
    """
    filter_status = request.args.get('status', '')
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = c.fetchall()

    query = """
        SELECT r.id, r.student_id, s.name, r.class, r.subject, r.status, r.requested_at, r.reviewed_at, r.admin_note
        FROM reattempt_requests r
        JOIN students s ON r.student_id = s.student_id
        WHERE 1=1
    """
    params = []
    if filter_status:
        query += " AND r.status = ?"
        params.append(filter_status)

    if assignments:
        where_clauses = []
        for a in assignments:
            cls = str(a['class']).strip()
            clean_cls = cls.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            subj = str(a['subject']).strip()
            where_clauses.append("((r.class=? OR r.class=? OR r.class LIKE ?) AND (LOWER(r.subject)=LOWER(?) OR r.subject LIKE ?))")
            params.extend([cls, clean_cls, f"%{clean_cls}%", subj, f"%{subj}%"])
        query += " AND (" + " OR ".join(where_clauses) + ")"

    query += " ORDER BY r.requested_at DESC"
    c.execute(query, params)
    requests = [dict(row) for row in c.fetchall()]
    c.execute("SELECT is_active FROM exam_control WHERE id=1")
    exam_row = c.fetchone()
    exam_is_active = bool(exam_row['is_active']) if exam_row else False
    conn.close()
    return jsonify({'requests': requests, 'exam_is_active': exam_is_active})

def process_coding_auto_submissions(student_ids, conn=None):
    """
    Auto-submit last-autosaved code drafts for students when their exam session expires or is ended.
    Evaluates un-submitted drafts and records 0 marks for unattempted questions, then updates results table.
    """
    if not student_ids:
        return
        
    close_db = False
    if conn is None:
        conn = get_db()
        close_db = True
        
    try:
        import json
        c = conn.cursor()
        for sid in student_ids:
            c.execute("SELECT * FROM students WHERE student_id=?", (sid,))
            student = c.fetchone()
            if not student:
                continue
                
            test_no = student['test_no'] or ''
            clean_cls = (student['class'] or '').replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            
            # Fetch all approved coding questions for this test
            c.execute("""SELECT id, total_marks, sample_input, sample_output FROM coding_questions
                         WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND (test_no=? OR topic=? OR test_no='') AND approved=1""",
                      (student['class'], f"%{clean_cls}%", student['subject'], test_no, test_no))
            questions = c.fetchall()
            
            if not questions:
                continue
                
            for q in questions:
                qid = q['id']
                qmarks = float(q['total_marks'] or 10.0)
                
                # Check if submission already exists
                c.execute("SELECT id FROM coding_submissions WHERE student_id=? AND question_id=?", (sid, qid))
                sub_row = c.fetchone()
                if sub_row:
                    continue
                    
                # Check if draft exists for this student + question
                c.execute("""SELECT language, code FROM coding_drafts 
                             WHERE student_id=? AND question_id=? AND code != '' 
                             ORDER BY updated_at DESC LIMIT 1""", (sid, qid))
                draft_row = c.fetchone()
                
                if draft_row and draft_row['code'].strip():
                    lang = draft_row['language']
                    code = draft_row['code']
                    
                    c.execute("""SELECT id, input, expected_output, is_hidden, marks_weight 
                                 FROM coding_test_cases WHERE question_id=? ORDER BY id""", (qid,))
                    tc_rows = [dict(r) for r in c.fetchall()]
                    if not tc_rows:
                        tc_rows = [{'id': 1, 'input': q['sample_input'] or '', 'expected_output': q['sample_output'] or '', 'is_hidden': 0, 'marks_weight': 1.0}]
                        
                    eval_res = evaluate_submission(code, lang, tc_rows, question_marks=qmarks, timeout_seconds=5)
                    
                    sanitized_results = []
                    for r in eval_res['results']:
                        is_h = bool(r.get('is_hidden'))
                        sanitized_results.append({
                            'test_case_id': r.get('test_case_id'),
                            'is_hidden': is_h,
                            'passed': bool(r.get('passed')),
                            'status': r.get('status'),
                            'input': '[Hidden]' if is_h else r.get('input'),
                            'expected_output': '[Hidden]' if is_h else r.get('expected_output'),
                            'actual_output': '[Hidden]' if is_h else r.get('actual_output'),
                            'execution_time_ms': round(float(r.get('execution_time_ms', 0)), 1),
                            'stderr': '[Hidden]' if is_h else r.get('stderr')
                        })
                        
                    c.execute("""INSERT INTO coding_submissions 
                                 (student_id, question_id, code, language, test_cases_passed, total_test_cases, marks_awarded, status, stdout_output, stderr_output, submitted_at)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)""",
                              (sid, qid, code, lang, eval_res['passed_count'], eval_res['total_test_cases'], eval_res['marks_awarded'], 'Auto-Submitted', json.dumps(sanitized_results), ''))
                else:
                    c.execute("""SELECT COUNT(*) FROM coding_test_cases WHERE question_id=?""", (qid,))
                    tc_cnt = c.fetchone()[0] or 1
                    c.execute("""INSERT INTO coding_submissions 
                                 (student_id, question_id, code, language, test_cases_passed, total_test_cases, marks_awarded, status, stdout_output, stderr_output, submitted_at)
                                 VALUES (?, ?, ?, ?, 0, ?, 0.0, 'Unattempted', '[]', '', CURRENT_TIMESTAMP)""",
                              (sid, qid, '', 'python', tc_cnt))
                              
            sync_student_coding_results(sid, conn=conn)
            
        conn.commit()
    except Exception as e:
        app.logger.error(f"Error in process_coding_auto_submissions: {e}")
    finally:
        if close_db:
            conn.close()

def auto_submit_expired_exams():
    """
    Auto-submit exams that have exceeded their allotted time.
    Handles per-session expiry and pause awareness, and falls back to global exam_control.
    """
    conn = get_db()
    c = conn.cursor()

    try:
        # ── 1. Per-Session Auto-Submit ────────────────────────────────────
        c.execute("SELECT * FROM exam_sessions WHERE status IN ('running', 'paused')")
        active_sessions = [dict(r) for r in c.fetchall()]

        for sess in active_sessions:
            if sess.get('status') == 'paused':
                continue

            rem_secs = get_session_remaining_seconds(sess)
            cls = sess.get('class', '')
            clean_cls = str(cls).replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            sec = sess.get('section', '')
            subj = sess.get('subject', '')
            tno = sess.get('test_no', '')

            scope_where = "(class=? OR class=? OR ? LIKE '%' || class || '%') AND LOWER(subject)=LOWER(?)"
            scope_params = [cls, clean_cls, cls, subj]
            if sec:
                scope_where += " AND LOWER(section)=LOWER(?)"
                scope_params.append(sec)
            if tno:
                scope_where += " AND (test_no=? OR test_no='' OR test_no IS NULL)"
                scope_params.append(tno)

            if rem_secs <= 0:
                c.execute("""UPDATE exam_sessions 
                             SET status='ended', end_time=CURRENT_TIMESTAMP, end_reason='time_expired' 
                             WHERE id=?""", (sess['id'],))
                
                c.execute(f"SELECT student_id FROM students WHERE status='In Progress' AND {scope_where}", scope_params)
                expired_sids = [r['student_id'] for r in c.fetchall()]
                
                c.execute(f"UPDATE students SET status='Submitted' WHERE status='In Progress' AND {scope_where}", scope_params)
                
                if expired_sids:
                    process_coding_auto_submissions(expired_sids, conn=conn)
            else:
                tot_paused_mins = int((sess.get('total_paused_secs') or 0) // 60)
                allowed_mins = int(sess.get('allotted_duration') or 60) + tot_paused_mins
                
                c.execute(f"""
                    SELECT student_id FROM students 
                    WHERE status='In Progress'
                      AND exam_started_at IS NOT NULL
                      AND {scope_where}
                      AND datetime(exam_started_at, ? || ' minutes') < datetime('now')
                """, tuple(scope_params + [str(allowed_mins)]))
                indiv_sids = [r['student_id'] for r in c.fetchall()]

                c.execute(f"""
                    UPDATE students 
                    SET status='Submitted'
                    WHERE status='In Progress'
                      AND exam_started_at IS NOT NULL
                      AND {scope_where}
                      AND datetime(exam_started_at, ? || ' minutes') < datetime('now')
                """, tuple(scope_params + [str(allowed_mins)]))
                
                if indiv_sids:
                    process_coding_auto_submissions(indiv_sids, conn=conn)

        # ── 2. Legacy Global Auto-Submit (Fallback) ───────────────────────
        c.execute("SELECT id, duration, is_active FROM exam_control WHERE id=1")
        exam = c.fetchone()
        if exam and exam['duration'] and exam['is_active']:
            duration = exam['duration']
            c.execute("""
                SELECT student_id FROM students 
                WHERE status='In Progress'
                AND exam_started_at IS NOT NULL
                AND datetime(exam_started_at, ? || ' minutes') < datetime('now')
            """, (str(duration),))
            glob_sids = [r['student_id'] for r in c.fetchall()]

            c.execute("""
                UPDATE students 
                SET status='Submitted'
                WHERE status='In Progress'
                AND exam_started_at IS NOT NULL
                AND datetime(exam_started_at, ? || ' minutes') < datetime('now')
            """, (str(duration),))

            if glob_sids:
                process_coding_auto_submissions(glob_sids, conn=conn)

        conn.commit()
    except Exception as e:
        app.logger.error(f"Error in auto_submit_expired_exams: {e}")
    finally:
        conn.close()

@app.route('/guidelines')
@student_required
def guidelines():
    # Mark that the student has not yet accepted guidelines for this session
    session['guidelines_accepted'] = False
    # Clear resume_mode after it has been displayed on the waiting page;
    # it should not persist beyond the guidelines screen.
    session.pop('resume_mode', None)
    return render_template('guidelines.html')

@app.route('/accept_guidelines', methods=['POST'])
@student_required
def accept_guidelines():
    session['guidelines_accepted'] = True
    if session.get('test_type') == 'Coding Test':
        return redirect(url_for('coding_exam_list'))
    return redirect(url_for('exam'))

# ── TWO-STAGE INSTRUCTIONS ROUTES ─────────────────────────────────────────────

@app.route('/exam/instructions/stage1')
@student_required
def instructions_stage1():
    """Render Stage 1: General CBT instructions and palette legend."""
    session['guidelines_accepted'] = False
    session.pop('resume_mode', None)
    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')
    return render_template('instructions_stage1.html', school_name=school_name, logo_path=logo_path)

@app.route('/exam/instructions/stage2')
@student_required
def instructions_stage2():
    """Render Stage 2: Test structure table, marking schemes, language selector, and declaration."""
    student_id = session.get('student_id')
    test_paper_id = session.get('test_paper_id')
    timing_mode = session.get('timing_mode', 'disabled')
    subject = session.get('subject', '')
    test_no = session.get('test_no', '')
    class_ = session.get('class', '')

    conn = get_db()
    c = conn.cursor()
    if not test_paper_id:
        clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        c.execute("""SELECT tp.id, tsc.timing_mode FROM test_papers tp
                     JOIN test_section_config tsc ON tp.id = tsc.test_paper_id
                     WHERE (tp.class=? OR tp.class LIKE ?) AND LOWER(tp.subject)=LOWER(?) AND tp.test_no=?
                     LIMIT 1""", (class_, f"%{clean_cls}%", subject, test_no))
        tp_row = c.fetchone()
        if tp_row:
            test_paper_id = tp_row['id']
            timing_mode = tp_row['timing_mode']
            session['test_paper_id'] = test_paper_id
            session['timing_mode'] = timing_mode

    sections = []
    if test_paper_id:
        c.execute("""SELECT id, section_name, subject, question_count, positive_marks, negative_marks, duration_minutes, order_index
                     FROM test_sections
                     WHERE test_paper_id=?
                     ORDER BY order_index ASC""", (test_paper_id,))
        sections = [dict(r) for r in c.fetchall()]

    conn.close()

    total_questions = sum(s.get('question_count', 0) for s in sections)
    if timing_mode == 'fixed':
        total_duration = sum(s.get('duration_minutes', 0) for s in sections)
    else:
        total_duration = int(get_setting('exam_duration', '60') or 60)

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')
    return render_template('instructions_stage2.html',
                           sections=sections,
                           timing_mode=timing_mode,
                           test_title=subject,
                           test_no=test_no,
                           total_questions=total_questions,
                           total_duration=total_duration,
                           school_name=school_name,
                           logo_path=logo_path)

@app.route('/exam/instructions/accept', methods=['POST'])
@student_required
def instructions_accept():
    """Accept declaration, save language preference, and start exam."""
    session['guidelines_accepted'] = True
    exam_lang = (request.form.get('exam_language') or 'English').strip()
    session['exam_language'] = exam_lang
    lang_code = 'hi' if 'hindi' in exam_lang.lower() or exam_lang.lower() == 'hi' else ('en' if 'english' in exam_lang.lower() or exam_lang.lower() == 'en' else exam_lang.lower())
    session['student_lang'] = lang_code
    return redirect(url_for('exam'))

@app.route('/admin/reattempt_request/<int:req_id>/<action>', methods=['POST'])
@admin_required
def handle_reattempt_request(req_id, action):
    if action not in ('approve', 'reject'):
        return jsonify({'status': 'error', 'message': 'Invalid action'}), 400

    admin_note = request.json.get('note', '') if request.is_json else ''

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT student_id, class, subject, status FROM reattempt_requests WHERE id=?", (req_id,))
    req = c.fetchone()
    if not req:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Request not found'}), 404
    if req['status'] != 'pending':
        conn.close()
        return jsonify({'status': 'error', 'message': 'Request already processed'}), 400

    student_id = req['student_id']
    class_ = req['class']
    subject = req['subject']

    if action == 'approve':
        c.execute("""DELETE FROM responses WHERE student_id=? AND question_id IN (
                     SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                  )""", (student_id, class_, subject))
        c.execute("DELETE FROM results WHERE student_id=? AND class=? AND LOWER(subject)=LOWER(?)",
                  (student_id, class_, subject))
        c.execute("""DELETE FROM shuffled_questions WHERE student_id=? AND question_id IN (
                     SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                  )""", (student_id, class_, subject))
        c.execute("UPDATE students SET status='Not Started', exam_started_at=NULL WHERE student_id=?", (student_id,))
        c.execute("UPDATE reattempt_requests SET status='approved', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                  (admin_note, req_id))
        conn.commit()
    else:
        c.execute("UPDATE reattempt_requests SET status='rejected', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                  (admin_note, req_id))
        conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'REATTEMPT_' + action.upper(), 'reattempt_requests', req_id, request.remote_addr)
    return jsonify({'status': 'success'})

@app.route('/admin/reattempt_requests/bulk_action', methods=['POST'])
@admin_required
def admin_bulk_reattempt_action():
    """
    Approve or reject all (or specified) pending reattempt requests as Admin.
    """
    data = request.get_json(silent=True) or {}
    action = data.get('action', '').strip().lower()
    if action not in ('approve', 'reject'):
        return jsonify({'status': 'error', 'message': 'Invalid action. Must be approve or reject.'}), 400

    admin_note = (data.get('note') or '').strip()
    request_ids = data.get('request_ids', [])

    conn = get_db()
    c = conn.cursor()

    if request_ids:
        placeholders = ','.join(['?'] * len(request_ids))
        c.execute(f"SELECT id, student_id, class, subject FROM reattempt_requests WHERE id IN ({placeholders}) AND status='pending'", request_ids)
    else:
        c.execute("SELECT id, student_id, class, subject FROM reattempt_requests WHERE status='pending'")
    
    pending_reqs = [dict(r) for r in c.fetchall()]

    if not pending_reqs:
        conn.close()
        return jsonify({'status': 'error', 'message': 'No pending requests found to process.'}), 400

    processed_count = 0
    for req in pending_reqs:
        req_id = req['id']
        student_id = req['student_id']
        class_ = req['class']
        subject = req['subject']

        if action == 'approve':
            c.execute("""DELETE FROM responses WHERE student_id=? AND question_id IN (
                         SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                      )""", (student_id, class_, subject))
            c.execute("DELETE FROM results WHERE student_id=? AND class=? AND LOWER(subject)=LOWER(?)",
                      (student_id, class_, subject))
            c.execute("""DELETE FROM shuffled_questions WHERE student_id=? AND question_id IN (
                         SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                      )""", (student_id, class_, subject))
            c.execute("UPDATE students SET status='Not Started', exam_started_at=NULL WHERE student_id=?", (student_id,))
            c.execute("UPDATE reattempt_requests SET status='approved', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                      (admin_note, req_id))
        else:
            c.execute("UPDATE reattempt_requests SET status='rejected', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                      (admin_note, req_id))
        processed_count += 1

    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', f'REATTEMPT_BULK_{action.upper()}', 'reattempt_requests', processed_count, request.remote_addr)
    past_action = 'approved' if action == 'approve' else 'rejected'
    return jsonify({
        'status': 'success',
        'count': processed_count,
        'message': f"Successfully {past_action} {processed_count} reattempt request(s)."
    })

@app.route('/teacher/reattempt_request/<int:req_id>/<action>', methods=['POST'])
@teacher_required
def teacher_handle_reattempt_request(req_id, action):
    """
    API endpoint for teachers to approve or reject student reattempt requests.
    """
    if action not in ('approve', 'reject'):
        return jsonify({'status': 'error', 'message': 'Invalid action'}), 400

    teacher_note = request.json.get('note', '') if request.is_json else ''
    teacher_id = session.get('teacher_id')
    teacher_name = session.get('teacher_name', 'Teacher')

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT student_id, class, subject, status FROM reattempt_requests WHERE id=?", (req_id,))
    req = c.fetchone()
    if not req:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Request not found'}), 404
    if req['status'] != 'pending':
        conn.close()
        return jsonify({'status': 'error', 'message': 'Request already processed'}), 400

    student_id = req['student_id']
    class_ = req['class']
    subject = req['subject']

    note_text = f"Teacher ({teacher_name}): {teacher_note}".strip() if teacher_note else f"Reviewed by {teacher_name}"

    if action == 'approve':
        c.execute("""DELETE FROM responses WHERE student_id=? AND question_id IN (
                     SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                  )""", (student_id, class_, subject))
        c.execute("DELETE FROM results WHERE student_id=? AND class=? AND LOWER(subject)=LOWER(?)",
                  (student_id, class_, subject))
        c.execute("""DELETE FROM shuffled_questions WHERE student_id=? AND question_id IN (
                     SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                  )""", (student_id, class_, subject))
        c.execute("UPDATE students SET status='Not Started', exam_started_at=NULL WHERE student_id=?", (student_id,))
        c.execute("UPDATE reattempt_requests SET status='approved', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                  (note_text, req_id))
        conn.commit()
    else:
        c.execute("UPDATE reattempt_requests SET status='rejected', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                  (note_text, req_id))
        conn.commit()

    conn.close()
    log_audit_event('teacher', str(teacher_id), 'REATTEMPT_' + action.upper(), 'reattempt_requests', req_id, request.remote_addr)
    return jsonify({'status': 'success'})

@app.route('/teacher/reattempt_requests/bulk_action', methods=['POST'])
@teacher_required
def teacher_bulk_reattempt_action():
    """
    Approve or reject all (or specified) authorized pending reattempt requests as Teacher.
    """
    data = request.get_json(silent=True) or {}
    action = data.get('action', '').strip().lower()
    if action not in ('approve', 'reject'):
        return jsonify({'status': 'error', 'message': 'Invalid action. Must be approve or reject.'}), 400

    teacher_note = (data.get('note') or '').strip()
    teacher_id = session.get('teacher_id')
    teacher_name = session.get('teacher_name', 'Teacher')
    request_ids = data.get('request_ids', [])

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = c.fetchall()

    query = """
        SELECT r.id, r.student_id, r.class, r.subject
        FROM reattempt_requests r
        WHERE r.status = 'pending'
    """
    params = []
    if request_ids:
        placeholders = ','.join(['?'] * len(request_ids))
        query += f" AND r.id IN ({placeholders})"
        params.extend(request_ids)

    if assignments:
        where_clauses = []
        for a in assignments:
            cls = str(a['class']).strip()
            clean_cls = cls.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
            subj = str(a['subject']).strip()
            where_clauses.append("((r.class=? OR r.class=? OR r.class LIKE ?) AND (LOWER(r.subject)=LOWER(?) OR r.subject LIKE ?))")
            params.extend([cls, clean_cls, f"%{clean_cls}%", subj, f"%{subj}%"])
        query += " AND (" + " OR ".join(where_clauses) + ")"

    c.execute(query, params)
    pending_reqs = [dict(r) for r in c.fetchall()]

    if not pending_reqs:
        conn.close()
        return jsonify({'status': 'error', 'message': 'No authorized pending requests found to process.'}), 400

    note_text = f"Teacher ({teacher_name}): {teacher_note}".strip() if teacher_note else f"Reviewed by {teacher_name}"
    processed_count = 0
    for req in pending_reqs:
        req_id = req['id']
        student_id = req['student_id']
        class_ = req['class']
        subject = req['subject']

        if action == 'approve':
            c.execute("""DELETE FROM responses WHERE student_id=? AND question_id IN (
                         SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                      )""", (student_id, class_, subject))
            c.execute("DELETE FROM results WHERE student_id=? AND class=? AND LOWER(subject)=LOWER(?)",
                      (student_id, class_, subject))
            c.execute("""DELETE FROM shuffled_questions WHERE student_id=? AND question_id IN (
                         SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                      )""", (student_id, class_, subject))
            c.execute("UPDATE students SET status='Not Started', exam_started_at=NULL WHERE student_id=?", (student_id,))
            c.execute("UPDATE reattempt_requests SET status='approved', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                      (note_text, req_id))
        else:
            c.execute("UPDATE reattempt_requests SET status='rejected', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE id=?",
                      (note_text, req_id))
        processed_count += 1

    conn.commit()
    conn.close()
    log_audit_event('teacher', str(teacher_id), f'REATTEMPT_BULK_{action.upper()}', 'reattempt_requests', processed_count, request.remote_addr)
    past_action = 'approved' if action == 'approve' else 'rejected'
    return jsonify({
        'status': 'success',
        'count': processed_count,
        'message': f"Successfully {past_action} {processed_count} reattempt request(s)."
    })

@app.route('/admin/student/<student_id>/allow_reattempt', methods=['POST'])
@admin_required
def admin_allow_student_reattempt(student_id):
    """
    Directly allow a student to reattempt the exam by resetting their responses and status.
    """
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT class, subject FROM students WHERE student_id=?", (student_id,))
    st = c.fetchone()
    if st:
        c.execute("""DELETE FROM responses WHERE student_id=? AND question_id IN (
                     SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                  )""", (student_id, st['class'], st['subject']))
        c.execute("DELETE FROM results WHERE student_id=? AND class=? AND LOWER(subject)=LOWER(?)",
                  (student_id, st['class'], st['subject']))
        c.execute("""DELETE FROM shuffled_questions WHERE student_id=? AND question_id IN (
                     SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
                  )""", (student_id, st['class'], st['subject']))
    else:
        c.execute("DELETE FROM responses WHERE student_id=?", (student_id,))
        c.execute("DELETE FROM results WHERE student_id=?", (student_id,))
        c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
    c.execute("UPDATE students SET status='Not Started', exam_started_at=NULL WHERE student_id=?", (student_id,))
    c.execute("UPDATE reattempt_requests SET status='approved', reviewed_at=CURRENT_TIMESTAMP, admin_note='Direct reset by Admin' WHERE student_id=? AND status='pending'", (student_id,))
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'REATTEMPT_DIRECT_ALLOW', 'students', student_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': 'Student reset and allowed to reattempt exam'})

@app.route('/teacher/student/<student_id>/allow_reattempt', methods=['POST'])
@teacher_required
def teacher_allow_student_reattempt(student_id):
    """
    Directly allow a student to reattempt the exam from the teacher panel.
    """
    teacher_id = session.get('teacher_id')
    teacher_name = session.get('teacher_name', 'Teacher')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT class, subject FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student not found'}), 404

    c.execute("""DELETE FROM responses WHERE student_id=? AND question_id IN (
                 SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
              )""", (student_id, student['class'], student['subject']))
    c.execute("DELETE FROM results WHERE student_id=? AND class=? AND LOWER(subject)=LOWER(?)",
              (student_id, student['class'], student['subject']))
    c.execute("""DELETE FROM shuffled_questions WHERE student_id=? AND question_id IN (
                 SELECT id FROM questions WHERE class=? AND LOWER(subject)=LOWER(?)
              )""", (student_id, student['class'], student['subject']))
    c.execute("UPDATE students SET status='Not Started', exam_started_at=NULL WHERE student_id=?", (student_id,))
    c.execute("UPDATE reattempt_requests SET status='approved', reviewed_at=CURRENT_TIMESTAMP, admin_note=? WHERE student_id=? AND status='pending'",
              (f"Direct reset by {teacher_name}", student_id))
    conn.commit()
    conn.close()
    log_audit_event('teacher', str(teacher_id), 'REATTEMPT_DIRECT_ALLOW', 'students', student_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': 'Student reset and allowed to reattempt exam'})

# ========================
# TEACHER ROUTES
# ========================

@app.route('/teacher', methods=['GET', 'POST'])
@app.route('/teacher/login', methods=['GET', 'POST'])
def teacher_login():
    """
    Handle teacher login.
    """
    if request.method == 'POST':
        mobile = request.form.get('mobile')
        password = request.form.get('password')
        
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM teachers WHERE mobile=? AND status='active'", (mobile,))
        teacher = c.fetchone()
        conn.close()
        
        if teacher and verify_password(password, teacher['password']):
            session['teacher_logged_in'] = True
            session['teacher_id'] = teacher['id']
            session['teacher_name'] = teacher['name']
            session['teacher_mobile'] = teacher['mobile']
            return redirect(url_for('teacher_dashboard'))
        else:
            return render_template('teacher_login.html', error="Invalid credentials")
    
    return render_template('teacher_login.html')

@app.route('/teacher/logout')
def teacher_logout():
    """
    Handle teacher logout.
    """
    session.pop('teacher_logged_in', None)
    session.pop('teacher_id', None)
    session.pop('teacher_name', None)
    session.pop('teacher_mobile', None)
    return redirect(url_for('teacher_login'))

def process_cbt_question_image(q, class_, section, subject, chapter, output_dir=None):
    """
    Unified automatic diagram & image insertion pipeline for CBT questions.
    Feature parity with descriptive paper generator:
      1. Schemdraw: Electrical circuits (resistors, switches, capacitors, RLC, logic gates)
      2. RDKit: 2D Chemical structures from SMILES or compound names
      3. NCERT Local Bank: Handcrafted & curriculum diagrams
      4. Wikimedia Commons: Live API diagram search and caching
      5. Fallback: Clean styled callout placeholder image
    Returns relative path within /static/ (e.g. 'images/generated/...' or 'images/ncert/...')
    """
    if not isinstance(q, dict):
        return None

    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not output_dir:
        output_dir = os.path.join(script_dir, 'static', 'images', 'generated')
    os.makedirs(output_dir, exist_ok=True)

    clean_cls = re.sub(r'[^a-zA-Z0-9]', '', str(class_ or ''))
    clean_sec = re.sub(r'[^a-zA-Z0-9]', '', str(section or ''))
    clean_sub = re.sub(r'[^a-zA-Z0-9]', '', str(subject or ''))
    uid = uuid.uuid4().hex[:8]

    # 1. ELECTRICAL CIRCUIT (Schemdraw)
    circuit_type = q.get('circuit_type') or q.get('params', {}).get('circuit_type')
    is_circuit_cat = str(q.get('image_category', '')).lower() == 'circuit'
    q_text_lower = str(q.get('question', '')).lower()
    has_circuit_words = any(w in q_text_lower for w in ['circuit diagram', 'given circuit', 'resistors in series', 'resistors in parallel', 'wheatstone bridge', 'schematic diagram'])

    if circuit_type or is_circuit_cat or (has_circuit_words and ('physics' in str(subject).lower() or 'science' in str(subject).lower())):
        c_type = circuit_type or ('parallel_resistors' if 'parallel' in q_text_lower else 'series_resistors')
        params = q.get('params') or {}
        filename = f"cbt_circuit_{clean_cls}_{clean_sec}_{clean_sub}_{uid}.png"
        out_path = os.path.join(output_dir, filename)
        try:
            res_c = image_handler.generate_circuit(c_type, params, out_path)
            if res_c and os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                return f"images/generated/{filename}"
        except Exception as e:
            app.logger.warning(f"Error generating CBT circuit diagram: {e}")

    # 2. CHEMISTRY MOLECULE STRUCTURE (RDKit)
    smiles_val = q.get('smiles') or q.get('compound_name') or q.get('params', {}).get('smiles') or q.get('params', {}).get('compound_name')
    if smiles_val and str(smiles_val).startswith('__smiles__'):
        smiles_val = smiles_val[10:]
    is_chem_cat = str(q.get('image_category', '')).lower() == 'chemistry'
    if smiles_val or is_chem_cat:
        chem_id = smiles_val or q.get('image_prompt') or q.get('image_keywords')
        if isinstance(chem_id, list) and chem_id:
            chem_id = chem_id[0]
        if chem_id:
            filename = f"cbt_chem_{clean_cls}_{clean_sec}_{clean_sub}_{uid}.png"
            out_path = os.path.join(output_dir, filename)
            try:
                res_m = image_handler.generate_molecule_structure(str(chem_id), out_path)
                if res_m and os.path.exists(res_m) and os.path.getsize(res_m) > 0:
                    return f"images/generated/{filename}"
            except Exception as e:
                app.logger.warning(f"Error generating CBT molecular structure: {e}")

    # 3. DIAGRAMS (NCERT Image Bank -> Wikimedia Commons API)
    kws = q.get('image_keywords') or []
    if isinstance(kws, str):
        kws = [kws]
    if not kws and q.get('image_prompt'):
        kws = [q['image_prompt']]

    # Infer keywords from question text if explicitly visual
    needs_img = q.get('needs_image') or bool(kws) or any(phrase in q_text_lower for phrase in ['in the given diagram', 'in the figure below', 'shown in the figure', 'refer to the diagram', 'label the diagram', 'identify the part in the diagram'])
    if needs_img and not kws:
        kws = [str(chapter or subject)]

    if kws:
        try:
            # 3a. Search Local NCERT Bank
            matched_bank = image_handler.find_from_bank(kws)
            if matched_bank and os.path.exists(matched_bank):
                matched_norm = matched_bank.replace('\\', '/')
                if 'static/' in matched_norm:
                    return matched_norm.split('static/', 1)[1]
                return matched_norm
        except Exception as e:
            app.logger.warning(f"Error searching image bank for CBT: {e}")

        try:
            # 3b. Fetch from Wikimedia Commons API
            matched_wiki = image_handler.fetch_and_save_wikimedia_diagram(kws, topic=chapter, class_=class_, subject=subject)
            if matched_wiki and os.path.exists(matched_wiki):
                matched_norm = matched_wiki.replace('\\', '/')
                if 'static/' in matched_norm:
                    return matched_norm.split('static/', 1)[1]
                return matched_norm
        except Exception as e:
            app.logger.warning(f"Error fetching Wikimedia diagram for CBT: {e}")

    # 4. FALLBACK CALLOUT PLACEHOLDER (Only if an image was genuinely needed/requested)
    if needs_img and kws:
        try:
            kw_text = ", ".join([str(k) for k in kws[:3]]) if isinstance(kws, list) else str(kws)
            filename = f"cbt_ph_{clean_cls}_{clean_sec}_{uid}.png"
            out_path = os.path.join(output_dir, filename)
            from PIL import Image, ImageDraw
            img = Image.new('RGB', (460, 90), color='#f8fafc')
            draw = ImageDraw.Draw(img)
            draw.rectangle([1, 1, 458, 88], outline='#c7d2fe', width=2)
            draw.rectangle([4, 4, 455, 85], fill='#eef2ff')
            draw.text((20, 22), '[ DIAGRAM PLACEHOLDER ]', fill='#3730a3')
            draw.text((20, 48), f"Reference: {kw_text[:45]}", fill='#6366f1')
            img.save(out_path, format='PNG')
            if os.path.exists(out_path):
                return f"images/generated/{filename}"
        except Exception as e:
            app.logger.warning(f"Error generating CBT placeholder image: {e}")

    return None


@app.route('/teacher/create_test', methods=['GET', 'POST'])
@teacher_required
@limiter.limit("13 per minute")
def teacher_create_test():
    if request.method == 'POST':
        board = request.form.get('board', 'CBSE').strip() or 'CBSE'
        class_ = request.form.get('class')
        subject = request.form.get('subject')
        chapter = request.form.get('chapter')
        num_questions = int(request.form.get('num_questions', 10))
        prompt = request.form.get('prompt', '')
        method = request.form.get('method', 'ai')  # 'ai' or 'upload'
        
        if method == 'upload':
            # Handle CSV upload
            file = request.files.get('csv_file')
            if file and file.filename.endswith('.csv'):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                # Import questions from CSV
                with open(filepath, 'r', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    conn = get_db()
                    c = conn.cursor()
                    for row in reader:
                        c.execute("""INSERT INTO questions 
                                   (class, subject, chapter, question, option_a, option_b, option_c, option_d, correct_answer, board)
                                   VALUES (?,?,?,?,?,?,?,?,?,?)""",
                                (class_, subject, chapter,
                                 row.get('question', ''),
                                 row.get('option a', row.get('option_a', '')),
                                 row.get('option b', row.get('option_b', '')),
                                 row.get('option c', row.get('option_c', '')),
                                 row.get('option d', row.get('option_d', '')),
                                 row.get('correct_answer', ''),
                                 board))
                    conn.commit()
                    conn.close()
                os.remove(filepath)
                return jsonify({'status': 'success', 'message': 'Questions uploaded successfully'})
        
        else:  # AI generation using batched Multi-Provider AI
            try:
                type_counts = [('MCQ (4 options)', num_questions)]
                questions, gen_err = generate_cbt_questions_in_batches(
                    class_, '', subject, chapter, 'AI_Test', prompt,
                    type_counts=type_counts, board=board, max_batch_size=10
                )
                if not questions:
                    return jsonify({'status': 'error', 'message': gen_err or 'Failed to generate questions'}), 500

                # Save to database (including smiles and content_type for multimodal rendering)
                conn = get_db()
                c = conn.cursor()
                for q in questions:
                    img_path = process_cbt_question_image(q, class_, '', subject, chapter)
                    c.execute("""INSERT INTO questions 
                               (class, subject, chapter, question_type, question,
                                option_a, option_b, option_c, option_d, correct_answer, image_path, board)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (class_, subject, chapter,
                             'MCQ',
                             q.get('question', ''),
                             q.get('option_a', ''),
                             q.get('option_b', ''),
                             q.get('option_c', ''),
                             q.get('option_d', ''),
                             q.get('correct_answer', ''),
                             img_path, board))
                conn.commit()
                conn.close()

                return jsonify({'status': 'success', 'message': f'{len(questions)} questions generated and saved!'})

            except urllib.error.HTTPError as e:
                err_body = e.read().decode('utf-8') if hasattr(e, 'read') else str(e)
                app.logger.error(f"Gemini HTTP error: {err_body}")
                return jsonify({'status': 'error', 'message': f'Gemini API error: {e.code} - {err_body[:200]}'}), 500
            except Exception as e:
                app.logger.error(f"Teacher create test error: {e}")
                return jsonify({'status': 'error', 'message': str(e)}), 500
    
    # GET request - show form
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(row) for row in c.fetchall()]
    class_subject_map = {}
    for a in assignments:
        cls = a['class']
        if cls not in class_subject_map:
            class_subject_map[cls] = []
        if a['subject'] not in class_subject_map[cls]:
            class_subject_map[cls].append(a['subject'])
    category = request.args.get('category', '').strip()
    is_coding = (category == 'coding_test')
    return render_template('teacher_create_test_v2.html', assignments=assignments, class_subject_map=class_subject_map, category=category, is_coding=is_coding)
@app.route('/teacher/dashboard')
@teacher_required
def teacher_dashboard():
    """
    Render the teacher dashboard with allotted class stats, all-classes list,
    and same-subject teacher test counts.
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    
    # Get teacher assignments
    c.execute("""SELECT class, subject, is_class_teacher 
                 FROM teacher_assignments 
                 WHERE teacher_id=?""", (teacher_id,))
    assignments = [dict(row) for row in c.fetchall()]
    
    # Get teacher info
    c.execute("SELECT * FROM teachers WHERE id=?", (teacher_id,))
    t_row = c.fetchone()
    teacher = dict(t_row) if t_row else {}
    
    # Compile Detailed Allotted Classes Data
    allotted_details = []
    total_assigned_students = 0
    total_my_tests_all = 0

    for a in assignments:
        cls = str(a['class'])
        subj = str(a['subject'])
        
        # Enrolled students count for class
        c.execute('SELECT COUNT(*) FROM students WHERE class=?', (cls,))
        st_count = c.fetchone()[0]
        total_assigned_students += st_count
        
        # Total tests conducted for class & subject
        c.execute('SELECT COUNT(DISTINCT test_no) FROM questions WHERE class=? AND subject=?', (cls, subj))
        q_tests = c.fetchone()[0]
        c.execute('SELECT COUNT(*) FROM test_papers WHERE class=? AND subject=?', (cls, subj))
        tp_tests = c.fetchone()[0]
        tot_tests = max(q_tests, tp_tests)
        
        # My tests created for class & subject
        c.execute('SELECT COUNT(*) FROM test_generation_history WHERE teacher_id=? AND class=? AND subject=?', (teacher_id, cls, subj))
        my_ai = c.fetchone()[0]
        c.execute('SELECT COUNT(*) FROM test_papers WHERE uploaded_by=? AND class=? AND subject=?', (str(teacher_id), cls, subj))
        my_up = c.fetchone()[0]
        my_tests_count = my_ai + my_up
        total_my_tests_all += my_tests_count
        
        # Same subject teachers breakdown (all teachers teaching this subject across school)
        c.execute('''SELECT DISTINCT t.id, t.name, t.mobile, t.picture 
                     FROM teacher_assignments ta 
                     JOIN teachers t ON ta.teacher_id = t.id 
                     WHERE ta.subject=?''', (subj,))
        s_teachers = [dict(r) for r in c.fetchall()]
        
        teachers_breakdown = []
        for st in s_teachers:
            tid = st['id']
            c.execute('SELECT COUNT(*) FROM test_generation_history WHERE teacher_id=? AND subject=?', (tid, subj))
            tai = c.fetchone()[0]
            c.execute('SELECT COUNT(*) FROM test_papers WHERE uploaded_by=? AND subject=?', (str(tid), subj))
            tup = c.fetchone()[0]
            st['tests_count'] = tai + tup
            st['is_me'] = (tid == teacher_id)
            teachers_breakdown.append(st)
            
        allotted_details.append({
            'class': cls,
            'subject': subj,
            'is_class_teacher': a['is_class_teacher'],
            'student_count': st_count,
            'total_tests': tot_tests,
            'my_tests': my_tests_count,
            'teachers_breakdown': teachers_breakdown
        })

    # Compile All Classes List across System
    c.execute('''
        SELECT class FROM students WHERE class IS NOT NULL AND class != ''
        UNION
        SELECT class FROM questions WHERE class IS NOT NULL AND class != ''
        UNION
        SELECT class FROM teacher_assignments WHERE class IS NOT NULL AND class != ''
    ''')
    class_names = sorted([r[0] for r in c.fetchall()], key=lambda x: int(x) if str(x).isdigit() else str(x))

    all_classes_data = []
    for cls in class_names:
        c.execute('SELECT COUNT(*) FROM students WHERE class=?', (str(cls),))
        st_cnt = c.fetchone()[0]
        
        c.execute('SELECT DISTINCT subject FROM questions WHERE class=? AND subject IS NOT NULL', (str(cls),))
        q_subjs = [r[0] for r in c.fetchall()]
        c.execute('SELECT DISTINCT subject FROM teacher_assignments WHERE class=? AND subject IS NOT NULL', (str(cls),))
        ta_subjs = [r[0] for r in c.fetchall()]
        all_subjs = sorted(list(set(q_subjs + ta_subjs)))
        
        c.execute('SELECT COUNT(DISTINCT test_no) FROM questions WHERE class=?', (str(cls),))
        q_tests = c.fetchone()[0]
        c.execute('SELECT COUNT(*) FROM test_papers WHERE class=?', (str(cls),))
        tp_tests = c.fetchone()[0]
        tot_tests = max(q_tests, tp_tests)
        
        all_classes_data.append({
            'class': str(cls),
            'student_count': st_cnt,
            'subjects': all_subjs,
            'total_tests': tot_tests
        })

    conn.close()
    return render_template('teacher_dashboard.html', 
                         assignments=assignments,
                         teacher=teacher,
                         allotted_details=allotted_details,
                         all_classes_data=all_classes_data,
                         total_assigned_students=total_assigned_students,
                         total_my_tests_all=total_my_tests_all)



@app.route('/teacher/students')
@teacher_required
def teacher_students():
    """
    Render the page showing teacher's assigned students with step-by-step Class -> Section -> Student navigation.
    """
    teacher_id = session.get('teacher_id')
    class_filter = request.args.get('class', '').strip()
    section_filter = request.args.get('section', '').strip()
    
    conn = get_db()
    c = conn.cursor()
    
    # 1. Get teacher assignments
    c.execute("""SELECT class, section, subject, is_class_teacher 
                 FROM teacher_assignments 
                 WHERE teacher_id=?""", (teacher_id,))
    assignments = [dict(row) for row in c.fetchall()]
    
    # Build assigned classes map
    assigned_classes_map = {}
    for a in assignments:
        cls = str(a['class'])
        if not cls: continue
        if cls not in assigned_classes_map:
            assigned_classes_map[cls] = {
                'class_name': cls,
                'sections': set(),
                'subjects': [],
                'is_class_teacher': False,
                'total_students': 0
            }
        if a.get('section'):
            assigned_classes_map[cls]['sections'].add(str(a['section']))
        if a.get('subject') and a['subject'] not in assigned_classes_map[cls]['subjects']:
            assigned_classes_map[cls]['subjects'].append(a['subject'])
        if a.get('is_class_teacher'):
            assigned_classes_map[cls]['is_class_teacher'] = True
            
    for cls, cdata in assigned_classes_map.items():
        c.execute("SELECT DISTINCT section FROM students WHERE class=? AND section IS NOT NULL AND section!=''", (cls,))
        for r in c.fetchall():
            cdata['sections'].add(r[0])
        cdata['sections'] = sorted(list(cdata['sections']))
        
        c.execute("SELECT COUNT(*) FROM students WHERE class=?", (cls,))
        cdata['total_students'] = c.fetchone()[0]
        
    assigned_classes_list = sorted(list(assigned_classes_map.values()), key=lambda x: int(x['class_name']) if x['class_name'].isdigit() else x['class_name'])
    
    # 2. Build Student Query
    query = """SELECT DISTINCT s.student_id, s.name, s.admission_no, s.class, s.section, 
                      s.dob, s.house, s.parents_name, s.picture
               FROM students s
               WHERE 1=1"""
    params = []
    
    if class_filter:
        query += " AND s.class = ?"
        params.append(class_filter)
        if section_filter:
            query += " AND s.section = ?"
            params.append(section_filter)
    else:
        all_t_classes = list(set([str(a['class']) for a in assignments if a.get('class')]))
        if all_t_classes:
            placeholders = ','.join(['?'] * len(all_t_classes))
            query += f" AND s.class IN ({placeholders})"
            params.extend(all_t_classes)
            
    query += " ORDER BY s.class, s.section, s.name"
    c.execute(query, params)
    raw_students = [dict(row) for row in c.fetchall()]
    
    # 3. Compute student test metrics
    students = []
    for st in raw_students:
        sid = st['student_id']
        c.execute("SELECT COUNT(*) FROM results WHERE student_id=?", (sid,))
        t_count = c.fetchone()[0]
        
        c.execute("SELECT AVG(percentage) FROM results WHERE student_id=? AND percentage IS NOT NULL", (sid,))
        avg_perc = c.fetchone()[0]
        avg_perc = round(avg_perc, 1) if avg_perc is not None else 0.0
        
        st['tests_appeared'] = t_count
        st['avg_percentage'] = avg_perc
        students.append(st)
        
    conn.close()
    
    return render_template('teacher_students.html', 
                         assigned_classes=assigned_classes_list,
                         students=students,
                         assignments=assignments,
                         class_filter=class_filter,
                         section_filter=section_filter)

@app.route('/api/student_performance/<student_id>')
@teacher_required
def api_student_performance(student_id):
    """API endpoint to get detailed student performance analysis and graph data."""
    conn = get_db()
    c = conn.cursor()
    
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student not found'}), 404
    student = dict(student)
    
    teacher_id = session.get('teacher_id')
    c.execute("SELECT class, subject, is_class_teacher FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]
    is_class_teacher = any(a['class'] == student['class'] and a['is_class_teacher'] for a in assignments)
    
    if is_class_teacher:
        c.execute("""
            SELECT id, subject, chapter, score, total_questions, percentage, test_date
            FROM results
            WHERE student_id=?
            ORDER BY test_date ASC
        """, (student_id,))
    else:
        teacher_subjects = [a['subject'] for a in assignments if a['class'] == student['class']]
        if teacher_subjects:
            placeholders = ','.join(['?'] * len(teacher_subjects))
            c.execute(f"""
                SELECT id, subject, chapter, score, total_questions, percentage, test_date
                FROM results
                WHERE student_id=? AND subject IN ({placeholders})
                ORDER BY test_date ASC
            """, [student_id] + teacher_subjects)
        else:
            c.execute("SELECT id, subject, chapter, score, total_questions, percentage, test_date FROM results WHERE 1=0")
            
    results_list = [dict(r) for r in c.fetchall()]
    
    total_tests = len(results_list)
    avg_score = round(sum(r['percentage'] or 0 for r in results_list) / total_tests, 1) if total_tests > 0 else 0.0
    max_score = round(max((r['percentage'] or 0 for r in results_list), default=0.0), 1)
    passed_tests = sum(1 for r in results_list if (r['percentage'] or 0) >= 40.0)
    pass_rate = round((passed_tests / total_tests) * 100, 1) if total_tests > 0 else 0.0
    
    subject_graphs = {}
    for r in results_list:
        subj = r['subject'] or 'General'
        if subj not in subject_graphs:
            subject_graphs[subj] = {'labels': [], 'scores': [], 'max_marks': []}
        t_label = r['chapter'] if r.get('chapter') else f"Test #{r['id']}"
        label = f"{t_label} ({r['test_date'][:10] if r.get('test_date') else ''})"
        subject_graphs[subj]['labels'].append(label)
        subject_graphs[subj]['scores'].append(round(r['percentage'] or 0.0, 1))
        subject_graphs[subj]['max_marks'].append(f"{r['score']}/{r['total_questions']}")
        
    conn.close()
    
    return jsonify({
        'status': 'success',
        'student': student,
        'stats': {
            'total_tests': total_tests,
            'avg_score': avg_score,
            'max_score': max_score,
            'passed_tests': passed_tests,
            'pass_rate': pass_rate
        },
        'results': results_list,
        'subject_graphs': subject_graphs
    })

@app.route('/teacher/student/<student_id>')
@teacher_required
def teacher_student_profile(student_id):
    """
    Render the profile page of a student (teacher view).
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    
    # Get student info
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return "Student not found", 404
    student = dict(student)
    log_audit_event('teacher', session.get('teacher_id') or 'teacher', 'STUDENT_RESULT_VIEW', 'students', student_id, request.remote_addr)
    
    # Get teacher assignments
    c.execute("SELECT class, subject, is_class_teacher FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(row) for row in c.fetchall()]
    
    # Check if teacher is class teacher for this student
    is_class_teacher = any(a['class'] == student['class'] and a['is_class_teacher'] for a in assignments)
    
    # Get test history with safe numeric values
    if is_class_teacher:
        # Class teacher sees all subjects
        c.execute("""
            SELECT 
                r.id, r.subject, r.chapter, r.score, r.test_date,
                COALESCE(NULLIF(r.total_questions, 0), 
                    (SELECT COUNT(*) FROM questions WHERE class = r.class AND subject = r.subject AND (r.chapter IS NULL OR r.chapter = '' OR test_no = r.chapter OR chapter = r.chapter))) AS total_questions,
                COALESCE(r.percentage,
                    CASE 
                        WHEN (SELECT COUNT(*) FROM questions WHERE class = r.class AND subject = r.subject AND (r.chapter IS NULL OR r.chapter = '' OR test_no = r.chapter OR chapter = r.chapter)) > 0 
                        THEN ROUND((r.score * 100.0 / (SELECT COUNT(*) FROM questions WHERE class = r.class AND subject = r.subject AND (r.chapter IS NULL OR r.chapter = '' OR test_no = r.chapter OR chapter = r.chapter))), 2)
                        ELSE 0.0 
                    END) AS percentage
            FROM results r
            WHERE r.student_id = ?
            ORDER BY r.test_date DESC
        """, (student_id,))
    else:
        # Subject teacher sees only their subjects
        teacher_subjects = [a['subject'] for a in assignments if a['class'] == student['class']]
        if teacher_subjects:
            placeholders = ','.join(['?'] * len(teacher_subjects))
            c.execute(f"""
                SELECT 
                    r.id, r.subject, r.chapter, r.score, r.test_date,
                    COALESCE(NULLIF(r.total_questions, 0), 
                        (SELECT COUNT(*) FROM questions WHERE class = r.class AND subject = r.subject AND (r.chapter IS NULL OR r.chapter = '' OR test_no = r.chapter OR chapter = r.chapter))) AS total_questions,
                    COALESCE(r.percentage,
                        CASE 
                            WHEN (SELECT COUNT(*) FROM questions WHERE class = r.class AND subject = r.subject AND (r.chapter IS NULL OR r.chapter = '' OR test_no = r.chapter OR chapter = r.chapter)) > 0 
                            THEN ROUND((r.score * 100.0 / (SELECT COUNT(*) FROM questions WHERE class = r.class AND subject = r.subject AND (r.chapter IS NULL OR r.chapter = '' OR test_no = r.chapter OR chapter = r.chapter))), 2)
                            ELSE 0.0 
                        END) AS percentage
                FROM results r
                WHERE r.student_id = ? AND r.subject IN ({placeholders})
                ORDER BY r.test_date DESC
            """, [student_id] + teacher_subjects)
        else:
            c.execute("""
                SELECT 
                    r.id, r.subject, r.chapter, r.score, r.test_date,
                    COALESCE(r.total_questions, 0) AS total_questions,
                    COALESCE(r.percentage, 0.0) AS percentage
                FROM results r
                WHERE 1=0
            """)
    
    test_history = [dict(row) for row in c.fetchall()]
    
    # Generate chart data for progress
    chart_data = {}
    for test in test_history:
        subj = test['subject']
        if subj not in chart_data:
            chart_data[subj] = {'dates': [], 'percentages': [], 'marks': []}
        chart_data[subj]['dates'].append(test['test_date'][:10] if test['test_date'] else '')
        chart_data[subj]['percentages'].append(test['percentage'] or 0)
        chart_data[subj]['marks'].append(f"{test['score']}/{test['total_questions']}")
    
    conn.close()
    
    return render_template('teacher_student_profile.html',
                         student=student,
                         test_history=test_history,
                         chart_data=chart_data,
                         is_class_teacher=is_class_teacher)

@app.route('/teacher/student/<student_id>/edit', methods=['POST'])
@teacher_required
def teacher_edit_student(student_id):
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()

    # Get student class
    c.execute("SELECT class FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student not found'}), 404

    # Both class teachers AND subject teachers in that class can edit
    c.execute("""SELECT id FROM teacher_assignments
                 WHERE teacher_id=? AND class=?""", (teacher_id, student['class']))
    assignment = c.fetchone()
    if not assignment:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Access denied'}), 403

    name = request.form.get('name')
    admission_no = request.form.get('admission_no')
    class_ = request.form.get('class')
    section = request.form.get('section')
    dob = request.form.get('dob')
    house = request.form.get('house')
    parents_name = request.form.get('parents_name')
    address = request.form.get('address')

    picture_path = save_student_picture(request.files.get('picture'), student_id)

    if picture_path:
        c.execute("""UPDATE students
                     SET name=?, admission_no=?, class=?, section=?, dob=?, house=?,
                         parents_name=?, address=?, picture=?
                     WHERE student_id=?""",
                 (name, admission_no, class_, section, dob, house, parents_name, address,
                  picture_path, student_id))
    else:
        c.execute("""UPDATE students
                     SET name=?, admission_no=?, class=?, section=?, dob=?, house=?,
                         parents_name=?, address=?
                     WHERE student_id=?""",
                 (name, admission_no, class_, section, dob, house, parents_name, address, student_id))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Student details updated'})

@app.route('/teacher/student/add', methods=['POST'])
@teacher_required
def teacher_add_student():
    teacher_id = session.get('teacher_id')
    name = request.form.get('name', '').strip()
    student_id = request.form.get('student_id', '').strip()
    class_ = request.form.get('class', '').strip()
    section = request.form.get('section', '').strip()
    subject = request.form.get('subject', '').strip()
    admission_no = request.form.get('admission_no', '').strip()
    dob = request.form.get('dob', '').strip()
    house = request.form.get('house', '').strip()
    parents_name = request.form.get('parents_name', '').strip()
    address = request.form.get('address', '').strip()

    if not name or not student_id or not class_:
        return jsonify({'status': 'error', 'message': 'Name, Student ID and Class are required'}), 400

    # Verify teacher has access to this class
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM teacher_assignments WHERE teacher_id=? AND class=?", (teacher_id, class_))
    if not c.fetchone():
        conn.close()
        return jsonify({'status': 'error', 'message': 'You are not assigned to this class'}), 403

    # Check duplicate
    c.execute("SELECT student_id FROM students WHERE student_id=?", (student_id,))
    if c.fetchone():
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student ID already exists'}), 400

    picture_path = save_student_picture(request.files.get('picture'), student_id)

    c.execute("""INSERT INTO students
                 (student_id, name, class, subject, section, admission_no, dob, house, parents_name, address, picture, status)
                 VALUES (?,?,?,?,?,?,?,?,?,?,?,'Not Started')""",
             (student_id, name, class_, subject, section, admission_no, dob, house, parents_name, address, picture_path))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': f'Student {name} added successfully'})

@app.route('/teacher/student/<student_id>/delete', methods=['DELETE'])
@teacher_required
def teacher_delete_student(student_id):
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT class FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student not found'}), 404
    c.execute("SELECT id FROM teacher_assignments WHERE teacher_id=? AND class=?", (teacher_id, student['class']))
    if not c.fetchone():
        conn.close()
        return jsonify({'status': 'error', 'message': 'Access denied'}), 403
    c.execute("DELETE FROM students WHERE student_id=?",           (student_id,))
    c.execute("DELETE FROM responses WHERE student_id=?",          (student_id,))
    c.execute("DELETE FROM results WHERE student_id=?",            (student_id,))
    c.execute("DELETE FROM shuffled_questions WHERE student_id=?", (student_id,))
    c.execute("DELETE FROM student_class_lock WHERE student_id=?", (student_id,))  # Bug #3 fix
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Student deleted'})

@app.route('/teacher/profile', methods=['GET', 'POST'])
@teacher_required
def teacher_profile():
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip()
        address = request.form.get('address', '').strip()
        new_password = request.form.get('new_password', '').strip()
        pic_path = save_teacher_picture(request.files.get('picture'), teacher_id)
        if new_password:
            hashed = hash_password(new_password)
            if pic_path:
                c.execute("UPDATE teachers SET name=?, email=?, address=?, password=?, picture=? WHERE id=?",
                         (name, email, address, hashed, pic_path, teacher_id))
            else:
                c.execute("UPDATE teachers SET name=?, email=?, address=?, password=? WHERE id=?",
                         (name, email, address, hashed, teacher_id))
        else:
            if pic_path:
                c.execute("UPDATE teachers SET name=?, email=?, address=?, picture=? WHERE id=?",
                         (name, email, address, pic_path, teacher_id))
            else:
                c.execute("UPDATE teachers SET name=?, email=?, address=? WHERE id=?",
                         (name, email, address, teacher_id))
        conn.commit()
        session['teacher_name'] = name
        conn.close()
        return jsonify({'status': 'success', 'message': 'Profile updated'})

    c.execute("SELECT * FROM teachers WHERE id=?", (teacher_id,))
    teacher = dict(c.fetchone())
    c.execute("SELECT class, section, subject, is_class_teacher FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('teacher_profile.html', teacher=teacher, assignments=assignments)

@app.route('/teacher/test/<int:result_id>/responses')
@teacher_required
def teacher_view_test_responses(result_id):
    conn = get_db()
    c = conn.cursor()
    
    # Get test result details
    c.execute("SELECT * FROM results WHERE id=?", (result_id,))
    result = c.fetchone()
    if not result:
        conn.close()
        return "Test result not found", 404
    result = dict(result)
    
    # Parse section_breakdown if present
    if result.get('section_breakdown'):
        try:
            result['parsed_breakdown'] = json.loads(result['section_breakdown'])
        except Exception:
            result['parsed_breakdown'] = []
    else:
        result['parsed_breakdown'] = []

    # Get student responses with full question data, prioritizing shuffled_questions
    c.execute("""
        SELECT q.id, q.question, q.option_a, q.option_b, q.option_c, q.option_d, 
               q.correct_answer, q.image_path, r.selected_option,
               q.section_id, ts.section_name
        FROM shuffled_questions sq
        JOIN questions q ON sq.question_id = q.id
        LEFT JOIN responses r ON q.id = r.question_id AND r.student_id = ?
        LEFT JOIN test_sections ts ON q.section_id = ts.id
        WHERE sq.student_id = ?
        ORDER BY sq.shuffled_index
    """, (result['student_id'], result['student_id']))
    rows = c.fetchall()

    if not rows:
        c.execute("""
            SELECT q.id, q.question, q.option_a, q.option_b, q.option_c, q.option_d, 
                   q.correct_answer, q.image_path, r.selected_option,
                   NULL as section_id, '' as section_name
            FROM questions q
            LEFT JOIN responses r ON q.id = r.question_id AND r.student_id = ?
            WHERE q.class = ? AND q.subject = ?
            ORDER BY q.id
        """, (result['student_id'], result['class'], result['subject']))
        rows = c.fetchall()
    conn.close()
    
    # Process questions to map correct_answer to a letter
    questions = []
    for row in rows:
        q = dict(row)
        selected = q['selected_option'] or ''
        correct_raw = q['correct_answer']
        correct_letter = None
        
        if correct_raw:
            correct_raw = correct_raw.strip()
            # If already a single letter, use it
            if len(correct_raw) == 1 and correct_raw.upper() in ('A','B','C','D'):
                correct_letter = correct_raw.upper()
            else:
                # Try to match against option texts or "option_x"
                if correct_raw == q['option_a'] or correct_raw.lower() == 'option_a':
                    correct_letter = 'A'
                elif correct_raw == q['option_b'] or correct_raw.lower() == 'option_b':
                    correct_letter = 'B'
                elif correct_raw == q['option_c'] or correct_raw.lower() == 'option_c':
                    correct_letter = 'C'
                elif correct_raw == q['option_d'] or correct_raw.lower() == 'option_d':
                    correct_letter = 'D'
        
        is_correct = False
        if selected and correct_letter:
            is_correct = (selected.strip().upper() == correct_letter)
        
        sec_name = ''
        try:
            sec_name = q['section_name'] or ''
        except Exception:
            pass

        questions.append({
            'id': q['id'],
            'question': q['question'],
            'section_name': sec_name,
            'options': {
                'A': q['option_a'],
                'B': q['option_b'],
                'C': q['option_c'],
                'D': q['option_d']
            },
            'image_path': q['image_path'],
            'correct': correct_letter or '?',
            'selected': selected,
            'is_correct': is_correct
        })
    
    return render_template('teacher_test_responses.html',
                         result=result,
                         questions=questions)

@app.route('/teacher/generate_student_report/<student_id>')
@teacher_required
def teacher_generate_student_report(student_id):
    conn = get_db()
    c = conn.cursor()
    
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = dict(c.fetchone())
    
    c.execute("""SELECT subject, chapter, score, total_questions, percentage, test_date
                 FROM results WHERE student_id=?
                 ORDER BY test_date DESC""", (student_id,))
    test_history = [dict(row) for row in c.fetchall()]
    
    conn.close()
    
    # Generate PDF (simplified version)
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial; margin: 20px; }}
            h1 {{ color: #2e7d32; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #2e7d32; color: white; }}
        </style>
    </head>
    <body>
        <h1>Student Progress Report</h1>
        <p><strong>Name:</strong> {student.get('name', 'N/A')}</p>
        <p><strong>Admission No:</strong> {student.get('admission_no', 'N/A')}</p>
        <p><strong>Class:</strong> {student.get('class', 'N/A')} - {student.get('section', 'N/A')}</p>
        <p><strong>House:</strong> {student.get('house', 'N/A')}</p>
        
        <h2>Test History</h2>
        <table>
            <tr>
                <th>Date</th>
                <th>Subject</th>
                <th>Test No</th>
                <th>Marks</th>
                <th>Percentage</th>
            </tr>
    """
    
    for test in test_history:
        html_content += f"""
            <tr>
                <td>{test.get('test_date', '')[:10]}</td>
                <td>{test.get('subject', 'N/A')}</td>
                <td>{test.get('chapter') or '—'}</td>
                <td>{test.get('score', 0)}/{test.get('total_questions', 0)}</td>
                <td>{test.get('percentage', 0):.1f}%</td>
            </tr>
        """
    
    html_content += """
        </table>
    </body>
    </html>
    """
    
    from weasyprint import HTML
    pdf = HTML(string=html_content).write_pdf()
    
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'attachment; filename=student_report_{student_id}.pdf'
    
    return response

# ========================
# ADMIN TEACHER MANAGEMENT
# ========================

@app.route('/admin/teachers')
@admin_required
def admin_teachers():
    return render_template('admin_teachers.html')

@app.route('/admin/teachers/data')
@admin_required
def admin_teachers_data():
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT t.id, t.name, t.mobile, t.email, t.status, t.created_at,
                        GROUP_CONCAT(DISTINCT ta.class || ' - ' || ta.subject) as assignments
                 FROM teachers t
                 LEFT JOIN teacher_assignments ta ON t.id = ta.teacher_id
                 GROUP BY t.id
                 ORDER BY t.name""")
    teachers = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify({'teachers': teachers})

@app.route('/admin/teacher/add', methods=['POST'])
@admin_required
def admin_add_teacher():
    name = request.form.get('name')
    mobile = request.form.get('mobile')
    password = request.form.get('password')
    email = request.form.get('email')
    address = request.form.get('address')
    
    conn = get_db()
    c = conn.cursor()
    
    # Check if mobile exists
    c.execute("SELECT id FROM teachers WHERE mobile=?", (mobile,))
    if c.fetchone():
        conn.close()
        return jsonify({'status': 'error', 'message': 'Mobile number already exists'}), 400
    
    hashed_pw = hash_password(password)
    c.execute("""INSERT INTO teachers (name, mobile, password, email, address)
                 VALUES (?,?,?,?,?)""",
             (name, mobile, hashed_pw, email, address))
    conn.commit()
    conn.close()
    
    return jsonify({'status': 'success', 'message': 'Teacher added successfully'})

@app.route('/admin/teacher/<int:teacher_id>/assign', methods=['POST'])
@admin_required
def admin_assign_teacher(teacher_id):
    class_ = request.form.get('class', '').strip()
    section = request.form.get('section', '').strip()
    subject = request.form.get('subject', '').strip()
    is_class_teacher = request.form.get('is_class_teacher') == 'true'

    if not class_ or not subject:
        return jsonify({'status': 'error', 'message': 'Class and subject are required'}), 400

    conn = get_db()
    c = conn.cursor()

    # Enforce 1:1 class teacher constraint
    if is_class_teacher:
        c.execute("""SELECT ta.id, t.name FROM teacher_assignments ta
                     JOIN teachers t ON ta.teacher_id = t.id
                     WHERE ta.class=? AND ta.section=? AND ta.is_class_teacher=1
                     AND ta.teacher_id != ?""",
                 (class_, section, teacher_id))
        existing_ct = c.fetchone()
        if existing_ct:
            conn.close()
            return jsonify({'status': 'error',
                           'message': f"Class {class_} {section} already has a Class Teacher: {existing_ct['name']}"}), 400

    # Check if this exact assignment already exists
    c.execute("""SELECT id FROM teacher_assignments
                 WHERE teacher_id=? AND class=? AND section=? AND subject=?""",
             (teacher_id, class_, section, subject))
    existing = c.fetchone()

    if existing:
        c.execute("""UPDATE teacher_assignments
                     SET is_class_teacher=?
                     WHERE id=?""", (1 if is_class_teacher else 0, existing['id']))
    else:
        c.execute("""INSERT INTO teacher_assignments (teacher_id, class, section, subject, is_class_teacher)
                     VALUES (?,?,?,?,?)""",
                 (teacher_id, class_, section, subject, 1 if is_class_teacher else 0))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Assignment saved successfully'})

@app.route('/admin/teacher/<int:teacher_id>/remove_assignment/<int:assign_id>', methods=['DELETE'])
@admin_required
def admin_remove_assignment(teacher_id, assign_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM teacher_assignments WHERE id=? AND teacher_id=?", (assign_id, teacher_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/admin/teacher/<int:teacher_id>', methods=['GET'])
@admin_required
def admin_get_teacher(teacher_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, name, mobile, email, address, picture, status FROM teachers WHERE id=?", (teacher_id,))
    teacher = c.fetchone()
    if not teacher:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Teacher not found'}), 404
    c.execute("SELECT id, class, section, subject, is_class_teacher FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify({'teacher': dict(teacher), 'assignments': assignments})

@app.route('/admin/teacher/<int:teacher_id>/update', methods=['POST'])
@admin_required
def admin_update_teacher(teacher_id):
    conn = get_db()
    c = conn.cursor()
    name = request.form.get('name', '').strip()
    mobile = request.form.get('mobile', '').strip()
    email = request.form.get('email', '').strip()
    address = request.form.get('address', '').strip()
    status = request.form.get('status', 'active').strip()
    new_password = request.form.get('password', '').strip()

    if not name:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Teacher name cannot be empty'}), 400

    if mobile:
        c.execute("SELECT id FROM teachers WHERE mobile=? AND id != ?", (mobile, teacher_id))
        if c.fetchone():
            conn.close()
            return jsonify({'status': 'error', 'message': f'Mobile number {mobile} is already registered to another teacher'}), 400

    pic_file = request.files.get('picture')
    pic_path = save_teacher_picture(pic_file, teacher_id)

    if new_password:
        hashed = hash_password(new_password)
        if pic_path:
            c.execute("UPDATE teachers SET name=?, mobile=?, email=?, address=?, status=?, password=?, picture=? WHERE id=?",
                     (name, mobile, email, address, status, hashed, pic_path, teacher_id))
        else:
            c.execute("UPDATE teachers SET name=?, mobile=?, email=?, address=?, status=?, password=? WHERE id=?",
                     (name, mobile, email, address, status, hashed, teacher_id))
    else:
        if pic_path:
            c.execute("UPDATE teachers SET name=?, mobile=?, email=?, address=?, status=?, picture=? WHERE id=?",
                     (name, mobile, email, address, status, pic_path, teacher_id))
        else:
            c.execute("UPDATE teachers SET name=?, mobile=?, email=?, address=?, status=? WHERE id=?",
                     (name, mobile, email, address, status, teacher_id))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Teacher updated successfully'})

@app.route('/admin/teacher/<int:teacher_id>/delete', methods=['POST', 'DELETE'])
@admin_required
def admin_delete_teacher(teacher_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    c.execute("DELETE FROM teachers WHERE id=?", (teacher_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Teacher removed successfully'})

@app.route('/admin/teacher/upload_csv', methods=['POST'])
@admin_required
def admin_upload_teachers_csv():
    file = request.files.get('file')
    if not file or not file.filename.endswith('.csv'):
        return jsonify({'status': 'error', 'message': 'Invalid file'}), 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    conn = get_db()
    c = conn.cursor()
    added = 0
    
    with open(filepath, 'r', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            name = row.get('name', '').strip()
            mobile = row.get('mobile', '').strip()
            password = row.get('password', 'teacher123')  # Default password
            email = row.get('email', '').strip()
            
            if not name or not mobile:
                continue
            
            # Check if exists
            c.execute("SELECT id FROM teachers WHERE mobile=?", (mobile,))
            if c.fetchone():
                continue
            
            hashed_pw = hash_password(password)
            c.execute("""INSERT INTO teachers (name, mobile, password, email)
                         VALUES (?,?,?,?)""",
                     (name, mobile, hashed_pw, email))
            added += 1
    
    conn.commit()
    conn.close()
    os.remove(filepath)
    
    return jsonify({'status': 'success', 'message': f'{added} teachers added successfully'})

# ===================================================
# TEST PAPER MODULE (CSV naming: Class_Subject_TestNo)
# ===================================================

@app.route('/admin/test_papers')
@admin_required
def admin_test_papers():
    return render_template('test_papers.html', role='admin')

@app.route('/teacher/test_papers')
@teacher_required
def teacher_test_papers():
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('test_papers.html', role='teacher', assignments=assignments)

@app.route('/api/test_papers')
def api_test_papers():
    role = request.args.get('role', 'admin')
    teacher_id = session.get('teacher_id') if role == 'teacher' else None
    conn = get_db()
    c = conn.cursor()
    if role == 'teacher' and teacher_id:
        c.execute("""SELECT ta.class, ta.subject FROM teacher_assignments ta WHERE ta.teacher_id=?""", (teacher_id,))
        assigns = c.fetchall()
        papers = []
        for a in assigns:
            c.execute("""SELECT tp.*, t.name as uploader_name FROM test_papers tp
                         LEFT JOIN teachers t ON tp.uploaded_by = t.id AND tp.uploader_type='teacher'
                         WHERE tp.class=? AND tp.subject=?
                         ORDER BY tp.created_at DESC""", (a['class'], a['subject']))
            papers.extend([dict(r) for r in c.fetchall()])
    else:
        c.execute("""SELECT tp.* FROM test_papers tp ORDER BY tp.created_at DESC""")
        papers = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'papers': papers})

@app.route('/api/test_papers/upload', methods=['POST'])
def upload_test_paper():
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401

    file = request.files.get('csv_file')
    if not file or not file.filename.endswith('.csv'):
        return jsonify({'status': 'error', 'message': 'CSV file required'}), 400

    # Parse filename convention: Class_Subject_TestNo.csv
    raw_name = os.path.splitext(file.filename)[0]
    parts = raw_name.split('_')
    if len(parts) < 3:
        return jsonify({'status': 'error',
                        'message': 'Filename must follow format: Class_Subject_TestNo (e.g. 10A_Math_Test01)'}), 400

    class_ = parts[0].strip()
    test_no = parts[-1].strip()
    subject = '_'.join(parts[1:-1]).strip()
    section = ''
    # If class has letter suffix treat it as section: e.g. 10A -> class=10, section=A
    if class_ and class_[-1].isalpha():
        section = class_[-1].upper()
        class_ = class_[:-1]

    uploader_type = 'admin' if session.get('admin_logged_in') else 'teacher'
    uploaded_by = 'admin' if uploader_type == 'admin' else str(session.get('teacher_id'))

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    # Import questions
    conn = get_db()
    c = conn.cursor()
    question_count = 0
    try:
        with open(filepath, 'r', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                parsed = parse_question_csv_row(row)
                if not parsed or not parsed['question']:
                    continue
                c.execute("""INSERT INTO questions
                             (class, subject, chapter, question, option_a, option_b, option_c, option_d, correct_answer)
                             VALUES (?,?,?,?,?,?,?,?,?)""",
                         (class_, subject, test_no, parsed['question'],
                          parsed['option_a'], parsed['option_b'], parsed['option_c'], parsed['option_d'],
                          parsed['correct_answer']))
                question_count += 1

        c.execute("""INSERT INTO test_papers
                     (filename, class, section, subject, test_no, uploaded_by, uploader_type, question_count)
                     VALUES (?,?,?,?,?,?,?,?)""",
                 (raw_name, class_, section, subject, test_no, uploaded_by, uploader_type, question_count))

        # Record into mcq_test_history table
        t_id = session.get('teacher_id')
        t_name = uploader_type.capitalize()
        if t_id:
            c.execute("SELECT name FROM teachers WHERE id=?", (t_id,))
            t_r = c.fetchone()
            if t_r: t_name = t_r['name']
        c.execute("""INSERT INTO mcq_test_history (teacher_id, teacher_name, class, section, subject, test_no, question_count)
                     VALUES (?,?,?,?,?,?,?)""",
                  (t_id, t_name, class_, section, subject, test_no, question_count))

        c.execute("""INSERT INTO test_generation_history
                     (teacher_id,class,section,subject,chapter,test_no,output_mode,
                      total_questions,mcq_count,assertion_count,very_short_count,
                      short_count,long_count,case_study_count,remark)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (uploaded_by, class_, section, subject, test_no, test_no, 'cbt',
                   question_count, question_count, 0, 0, 0, 0, 0, 'Test Paper CSV Upload'))
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        os.remove(filepath)
        return jsonify({'status': 'error', 'message': str(e)}), 500
    conn.close()
    os.remove(filepath)
    return jsonify({'status': 'success',
                    'message': f'Uploaded {question_count} questions for {class_}{section} {subject} - {test_no}'})

@app.route('/api/test_papers/<int:paper_id>/toggle', methods=['POST'])
def toggle_test_paper(paper_id):
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, filename, test_no, class, subject, is_active FROM test_papers WHERE id=?", (paper_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Not found'}), 404
    new_status = 0 if row['is_active'] else 1
    c.execute("UPDATE test_papers SET is_active=? WHERE id=?", (new_status, paper_id))
    conn.commit()
    conn.close()
    
    status_label = "Active (visible to students)" if new_status else "Inactive (hidden from students)"
    msg = f"Test Paper '{row['filename']}' is now {status_label}."
    return jsonify({'status': 'success', 'is_active': new_status, 'message': msg})

@app.route('/api/test_papers/<int:paper_id>/delete', methods=['DELETE', 'POST'])
def delete_test_paper(paper_id):
    """Delete a CBT test paper and all its associated questions from the database."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM test_papers WHERE id=?", (paper_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Test paper not found'}), 404
    
    paper = dict(row)
    # Check teacher authorization
    if session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        t_id = str(session.get('teacher_id'))
        if str(paper.get('uploaded_by')) != t_id:
            c.execute("SELECT 1 FROM teacher_assignments WHERE teacher_id=? AND class=? AND subject=?",
                      (t_id, paper['class'], paper['subject']))
            if not c.fetchone():
                conn.close()
                return jsonify({'status': 'error', 'message': 'Not authorized to delete this test paper'}), 403
    
    # Delete associated questions
    c.execute("""DELETE FROM questions 
                 WHERE (class=? OR class=?) 
                   AND (LOWER(subject)=LOWER(?)) 
                   AND (test_no=? OR chapter=?)""",
              (paper['class'],
               paper['class'].replace('th','').replace('st','').replace('nd','').replace('rd',''),
               paper['subject'], paper['test_no'], paper['test_no']))
    del_questions = c.rowcount

    # Delete test paper entry
    c.execute("DELETE FROM test_papers WHERE id=?", (paper_id,))
    
    # Also clean up from mcq_test_history and test_generation_history
    c.execute("""DELETE FROM mcq_test_history WHERE class=? AND LOWER(subject)=LOWER(?) AND test_no=?""",
              (paper['class'], paper['subject'], paper['test_no']))
    c.execute("""DELETE FROM test_generation_history WHERE class=? AND LOWER(subject)=LOWER(?) AND test_no=?""",
              (paper['class'], paper['subject'], paper['test_no']))
    
    conn.commit()
    conn.close()
    
    u_type = 'teacher' if session.get('teacher_logged_in') else 'admin'
    u_id = session.get('teacher_id') if session.get('teacher_logged_in') else 0
    log_audit_event(u_type, u_id, 'CBT_PAPER_DELETE', 'test_papers', paper_id, request.remote_addr)

    return jsonify({
        'status': 'success',
        'message': f"Test Paper '{paper.get('filename')}' and {del_questions} CBT questions deleted successfully."
    })

@app.route('/api/test_papers/<int:paper_id>/pdf')
def test_paper_pdf(paper_id):
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return redirect(url_for('teacher_login'))
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM test_papers WHERE id=?", (paper_id,))
    paper = c.fetchone()
    if not paper:
        conn.close()
        return "Not found", 404
    paper = dict(paper)

    # Robust multi-fallback query for test paper PDF questions
    c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                        correct_answer, image_path,
                        CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                 FROM questions
                 WHERE (class=? OR class=? OR class LIKE ?)
                   AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                   AND (test_no=? OR chapter=? OR test_no LIKE ? OR chapter LIKE ?)
                 ORDER BY id""",
             (paper['class'],
              paper['class'].replace('th','').replace('st','').replace('nd','').replace('rd',''),
              f"%{paper['class']}%",
              paper['subject'], f"%{paper['subject']}%",
              paper['test_no'], paper['test_no'],
              f"%{paper['test_no']}%", f"%{paper['test_no']}%"))
    questions = [dict(r) for r in c.fetchall()]

    if not questions:
        clean_class = paper['class'].replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                            correct_answer, image_path,
                            CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                     FROM questions
                     WHERE (LOWER(class)=LOWER(?) OR LOWER(class)=LOWER(?) OR class LIKE ?)
                       AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                     ORDER BY id""",
                 (paper['class'], clean_class, f"%{clean_class}%", paper['subject'], f"%{paper['subject']}%"))
        questions = [dict(r) for r in c.fetchall()]

    # Fallback: Parse paper['filename'] (e.g. 6A_Computer_Test01) if initial query matches 0 questions
    if not questions and paper.get('filename'):
        parts = paper['filename'].split('_')
        if len(parts) >= 3:
            raw_cls = parts[0]
            t_no = parts[-1]
            subj = '_'.join(parts[1:-1]).replace('_', ' ')
            m = re.match(r'^(\d+)([A-Za-z]*)', raw_cls)
            if m:
                extracted_cls = m.group(1)
                c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                                    correct_answer, image_path,
                                    CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                             FROM questions
                             WHERE (class=? OR class LIKE ?)
                               AND (LOWER(subject)=LOWER(?) OR LOWER(subject)=LOWER(?))
                             ORDER BY id""",
                         (extracted_cls, f"%{extracted_cls}%", subj, paper.get('subject','')))
                questions = [dict(r) for r in c.fetchall()]

    if not questions:
        c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                            correct_answer, image_path,
                            CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                     FROM questions
                     ORDER BY id LIMIT 50""")
        questions = [dict(r) for r in c.fetchall()]

    conn.close()

    for q in questions:
        q_type = str(q.get('question_type') or '').strip()
        if not q_type or q_type.upper() in ['MCQ', 'CHOICE', 'OBJECTIVE']:
            q['question_type'] = 'MCQ'

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path   = get_setting('logo_path', '')
    logo_base64 = None
    if logo_path:
        abs_path = os.path.join(app.static_folder, logo_path)
        if os.path.exists(abs_path):
            with open(abs_path, "rb") as f:
                logo_bytes = f.read()
                ext  = os.path.splitext(abs_path)[1].lower()
                mime = "image/png" if ext == ".png" else "image/jpeg"
                logo_base64 = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode()}"

    # Determine display chapter name for paper PDF
    unique_ch = []
    for q in questions:
        ch_val = str(q.get('chapter') or '').strip()
        if ch_val and ch_val.lower() != str(paper['test_no']).lower() and ch_val not in unique_ch:
            unique_ch.append(ch_val)
    paper_chapter = ", ".join(unique_ch) if unique_ch else paper['test_no']

    # Use v2 template (handles all question types properly)
    rendered = render_template('question_paper_template_v2.html',
        school_name=school_name,
        logo_base64=logo_base64,
        class_name=f"{paper['class']}{paper['section']}",
        subject=paper['subject'],
        chapter=paper_chapter,
        test_no=paper['test_no'],
        questions=questions,
        total_marks=len(questions),
        date=datetime.datetime.now().strftime('%d/%m/%Y'))

    if request.args.get('format') == 'download_pdf':
        pdf = HTML(string=rendered, base_url=request.base_url).write_pdf()
        resp = make_response(pdf)
        resp.headers['Content-Type']        = 'application/pdf'
        resp.headers['Content-Disposition'] = f'attachment; filename={paper["filename"]}.pdf'
        return resp

    # Default: Return HTML Printable Preview (opens browser Print / Save as PDF dialog)
    return rendered

# ===================================================
# TEACHER MONITORING (own classes only)
# ===================================================

@app.route('/teacher/monitoring')
@teacher_required
def teacher_monitoring():
    """
    Render the teacher monitoring page for their assigned students.
    """
    return render_template('teacher_monitoring.html')

@app.route('/teacher/monitoring/data')
@teacher_required
def teacher_monitoring_data():
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    # Get teacher's assigned classes
    c.execute("SELECT DISTINCT class FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    classes = [r['class'] for r in c.fetchall()]
    if not classes:
        conn.close()
        return jsonify({'students': [], 'exam_is_active': False})

    placeholders = ','.join(['?'] * len(classes))
    c.execute(f"""SELECT s.student_id, s.name, s.class, s.subject, s.section, s.picture,
                         s.status, s.exam_started_at, s.current_section,
                         (SELECT COUNT(*) FROM responses r WHERE r.student_id=s.student_id) as answered
                  FROM students s
                  WHERE s.class IN ({placeholders})
                  ORDER BY s.class, s.name""", classes)
    students = [dict(r) for r in c.fetchall()]
    c.execute("SELECT is_active FROM exam_control WHERE id=1")
    exam_row = c.fetchone()
    conn.close()
    return jsonify({'students': students, 'exam_is_active': bool(exam_row['is_active']) if exam_row else False})

# ===================================================
# CLASS REPORT (Excel + PDF)
# ===================================================

@app.route('/admin/class_report')
@admin_required
def admin_class_report():
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT DISTINCT class FROM students WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class FROM results WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class FROM questions WHERE class IS NOT NULL AND class != ''
                 UNION
                 SELECT DISTINCT class FROM test_papers WHERE class IS NOT NULL AND class != ''""")
    classes = sorted(list(set(r['class'].strip() for r in c.fetchall() if r['class'] and r['class'].strip())))
    conn.close()
    return render_template('class_report.html', role='admin', classes=classes)

@app.route('/teacher/class_report')
@teacher_required
def teacher_class_report():
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, section, subject, is_class_teacher FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('class_report.html', role='teacher', assignments=assignments)

@app.route('/api/class_report/generate', methods=['POST'])
def generate_class_report():
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error'}), 401

    data = request.get_json(silent=True) or request.form or {}
    class_       = data.get('class', '').strip()
    section      = data.get('section', '').strip()
    subject      = data.get('subject', '').strip()
    date_from    = data.get('date_from', '').strip()
    date_to      = data.get('date_to', '').strip()
    export_format = data.get('format', 'excel')
    sort_by      = data.get('sort_by', 'name')

    if not class_:
        return jsonify({'status': 'error', 'message': 'Please enter a class to generate the report.'}), 400

    clean_cls = class_.replace('th','').replace('st','').replace('nd','').replace('rd','').strip()

    conn = get_db()
    c = conn.cursor()

    # 1. Fetch students from students table
    stu_query = """SELECT student_id, name, admission_no, class, section, house, parents_name, picture
                   FROM students WHERE (class=? OR class=? OR class LIKE ?)"""
    stu_params = [class_, clean_cls, f"%{clean_cls}%"]
    if section:
        stu_query += " AND (section=? OR section IS NULL OR section='')"
        stu_params.append(section)

    c.execute(stu_query, stu_params)
    db_students = [dict(r) for r in c.fetchall()]

    # 2. Fetch distinct student_ids from results table for this class
    res_query = """SELECT DISTINCT student_id, name, class, section FROM results
                   WHERE (class=? OR class=? OR class LIKE ?)"""
    res_params = [class_, clean_cls, f"%{clean_cls}%"]
    if section:
        res_query += " AND (section=? OR section IS NULL OR section='')"
        res_params.append(section)

    c.execute(res_query, res_params)
    res_students = [dict(r) for r in c.fetchall()]

    # Combine and deduplicate strictly by student_id (or name if student_id blank)
    student_map = {}
    for st in db_students:
        key = (st.get('student_id') or st.get('name') or '').strip().lower()
        if key and key not in student_map:
            st['admission_no'] = st.get('admission_no') or st.get('student_id') or ''
            student_map[key] = st

    for r in res_students:
        key = (r.get('student_id') or r.get('name') or '').strip().lower()
        if key and key not in student_map:
            student_map[key] = {
                'student_id': r.get('student_id'),
                'name': r.get('name') or key,
                'admission_no': r.get('student_id') or '',
                'class': r.get('class') or class_,
                'section': r.get('section') or section,
                'house': '',
                'parents_name': '',
                'picture': ''
            }

    students = list(student_map.values())

    # ── AB MARKING: Determine if the selected test window has expired ──────────
    # AB = student never submitted any result for this specific test AND the
    # scheduled exam window has definitively closed. Only applies when a specific
    # test_no is requested (not "All Tests").
    test_no = data.get('test_no', '').strip()
    is_test_window_expired = False
    if test_no and subject:
        c.execute("""
            SELECT status, end_time FROM scheduled_tests
            WHERE test_no=? AND (class=? OR class=? OR class LIKE ?)
              AND LOWER(subject)=LOWER(?) AND status != 'cancelled'
            ORDER BY CASE WHEN status='expired' THEN 0
                          WHEN status='active'  THEN 1
                          ELSE 2 END, id DESC LIMIT 1""",
            (test_no, class_, clean_cls, f"%{clean_cls}%", subject))
        _sch = c.fetchone()
        if _sch:
            if _sch['status'] == 'expired':
                is_test_window_expired = True
            elif _sch['status'] == 'active' and _sch['end_time']:
                # Edge case: status not yet updated but end_time has passed
                try:
                    _end_dt = datetime.datetime.fromisoformat(
                        str(_sch['end_time']).replace('T', ' '))
                    if _end_dt < datetime.datetime.now():
                        is_test_window_expired = True
                except Exception:
                    pass
    # ── END AB MARKING SETUP ──────────────────────────────────────────────────

    report_data = []
    for st in students:
        result_query  = """SELECT subject, chapter, score, total_questions, percentage, test_date, section_breakdown
                           FROM results WHERE student_id=?"""
        result_params = [st['student_id']]
        if subject:
            result_query += " AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)"
            result_params.extend([subject, f"%{subject}%"])
        if test_no:
            result_query += " AND (chapter=? OR chapter LIKE ?)"
            result_params.extend([test_no, f"%{test_no}%"])
        if date_from:
            result_query += " AND DATE(test_date)>=?"
            result_params.append(date_from)
        if date_to:
            result_query += " AND DATE(test_date)<=?"
            result_params.append(date_to)
        result_query += " ORDER BY test_date DESC"
        c.execute(result_query, result_params)
        results    = [dict(r) for r in c.fetchall()]
        for r in results:
            bd = r.get('section_breakdown')
            if bd and isinstance(bd, str):
                try:
                    r['parsed_breakdown'] = json.loads(bd)
                except Exception:
                    r['parsed_breakdown'] = []
            elif isinstance(bd, list):
                r['parsed_breakdown'] = bd
            else:
                r['parsed_breakdown'] = []
        total_score      = sum(r['score'] or 0 for r in results)
        total_possible   = sum(r['total_questions'] or 0 for r in results)
        avg_pct          = round(sum(r['percentage'] or 0 for r in results) / len(results), 1) if results else 0
        st['results']        = results
        st['avg_percentage'] = avg_pct
        st['total_tests']    = len(results)
        st['total_score']    = total_score
        st['total_possible'] = total_possible
        # ── AB MARKING: never-attempted student for an expired test ──────────
        if len(results) == 0 and is_test_window_expired:
            st['score_ratio'] = 'AB'
            st['is_ab']       = True
        elif total_possible:
            st['score_ratio'] = f"{total_score}/{total_possible}"
            st['is_ab']       = False
        else:
            st['score_ratio'] = '0/0'
            st['is_ab']       = False
        report_data.append(st)

    school_name = get_setting('school_name', 'RRB Group of Schools')
    conn.close()  # single close here

    if sort_by == 'percentage':
        report_data.sort(key=lambda x: x['avg_percentage'], reverse=True)
    else:
        report_data.sort(key=lambda x: x['name'])

    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    label     = f"{class_}{section}_{subject or 'All'}_{timestamp}"

    # Bug #5 Fix: json preview no longer calls conn.close() again
    if export_format == 'json':
        return jsonify({'status': 'success', 'students': report_data})

    elif export_format == 'excel':
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Class Report"

        # Header
        ws.merge_cells('A1:H1')
        ws['A1'] = school_name
        ws['A1'].font = openpyxl.styles.Font(bold=True, size=16)
        ws['A1'].alignment = openpyxl.styles.Alignment(horizontal='center')

        ws.merge_cells('A2:H2')
        ws['A2'] = f"Class: {class_}{section}  |  Subject: {subject or 'All'}  |  Test No: {test_no or 'All'}  |  Generated: {datetime.datetime.now().strftime('%d-%m-%Y')}"
        ws['A2'].font = openpyxl.styles.Font(bold=True, size=12)
        ws['A2'].alignment = openpyxl.styles.Alignment(horizontal='center')
        ws.append([])

        headers = ['#', 'Admission No / ID', 'Name', 'Class', 'Section', 'House', 'Total Tests', 'Score']
        ws.append(headers)
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.font = openpyxl.styles.Font(bold=True, color='FFFFFF')
            cell.fill = openpyxl.styles.PatternFill(start_color='1b5e20', end_color='1b5e20', fill_type='solid')
            cell.alignment = openpyxl.styles.Alignment(horizontal='center')

        for i, st in enumerate(report_data, 1):
            row_num = i + 4  # header is rows 1-4
            ws.append([i, st.get('admission_no') or st.get('student_id') or '', st['name'], st['class'],
                       st.get('section') or '', st.get('house') or '',
                       st['total_tests'], st.get('score_ratio', '0/0')])
            # Color AB score cell red/orange so it stands out in Excel
            if st.get('is_ab'):
                score_cell = ws.cell(row=row_num, column=8)
                score_cell.font = openpyxl.styles.Font(bold=True, color='FFFFFF')
                score_cell.fill = openpyxl.styles.PatternFill(start_color='c0392b', end_color='c0392b', fill_type='solid')
                score_cell.alignment = openpyxl.styles.Alignment(horizontal='center')

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(len(str(col[0].value or '')), 12) + 2

        filepath = os.path.join(app.config['EXPORT_FOLDER'], f"class_report_{label}.xlsx")
        wb.save(filepath)
        return send_file(filepath, as_attachment=True, download_name=f"class_report_{label}.xlsx")

    else:  # PDF / Printable Preview
        rows_html = ''
        for i, st in enumerate(report_data, 1):
            _is_ab = st.get('is_ab', False)
            _score_html = (
                '<strong style="color:#c0392b;background:#fde8e8;padding:2px 8px;border-radius:4px;">AB</strong>'
                if _is_ab
                else f'<strong>{st.get("score_ratio", "0/0")}</strong>'
            )
            _adm_no = st.get('admission_no') or st.get('student_id') or '—'
            _house = st.get('house') or '—'
            rows_html += f"""<tr>
                <td>{i}</td>
                <td>{_adm_no}</td>
                <td><strong>{st['name']}</strong></td>
                <td>{st['class']}{st.get('section') or ''}</td>
                <td>{_house}</td>
                <td>{st['total_tests']}</td>
                <td>{_score_html}</td>
            </tr>"""

        html = f"""<!DOCTYPE html><html><head>
        <meta charset="UTF-8">
        <title>Class Report - {school_name}</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
        <style>
            * {{ box-sizing: border-box; }}
            body {{ font-family: Arial, Helvetica, sans-serif; margin: 15mm; font-size: 10pt; color: #1e293b; background: #fff; line-height: 1.5; }}
            .print-toolbar {{
                position: sticky; top: 0; z-index: 9999;
                background: #1b5e20; color: #ffffff;
                padding: 10px 20px; margin: -15mm -15mm 20px -15mm;
                display: flex; align-items: center; justify-content: space-between; gap: 12px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            }}
            .print-toolbar .title {{ font-weight: 700; font-size: 1rem; display: flex; align-items: center; gap: 8px; }}
            .print-toolbar .actions {{ display: flex; align-items: center; gap: 10px; }}
            .print-btn {{
                background: #ffffff; color: #1b5e20; border: none; padding: 8px 16px; border-radius: 6px;
                font-weight: 700; font-size: 0.85rem; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; text-decoration: none;
            }}
            .print-btn:hover {{ background: #e8f5e9; }}
            .header-box {{ text-align: center; border-bottom: 2.5pt double #1b5e20; padding-bottom: 12px; margin-bottom: 16px; }}
            .school-title {{ font-size: 18pt; font-weight: 800; color: #1b5e20; }}
            .sub-title {{ font-size: 11pt; color: #475569; margin-top: 4px; font-weight: 600; }}
            .meta-bar {{ display: flex; justify-content: space-between; background: #f0fdf4; border: 1px solid #bbf7d0; padding: 8px 16px; border-radius: 8px; font-size: 9.5pt; margin-bottom: 16px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th {{ background: #1b5e20; color: white; padding: 8px 10px; text-align: left; font-size: 9.5pt; font-weight: 700; }}
            td {{ border-bottom: 1px solid #e2e8f0; padding: 8px 10px; font-size: 9pt; }}
            tr:nth-child(even) td {{ background: #f8fafc; }}
            .footer {{ margin-top: 30px; text-align: center; font-size: 9pt; color: #64748b; border-top: 1px solid #e2e8f0; padding-top: 10px; }}
            @media print {{
                body {{ margin: 10mm; }}
                .print-toolbar {{ display: none !important; }}
                th {{ -webkit-print-color-adjust: exact; }}
            }}
        </style></head><body>
        <div class="print-toolbar">
            <div class="title"><i class="fas fa-chart-bar"></i> Class Performance Report</div>
            <div class="actions">
                <button onclick="window.print()" class="print-btn"><i class="fas fa-print"></i> Print / Save as PDF</button>
            </div>
        </div>
        <div class="header-box">
            <div class="school-title">{school_name}</div>
            <div class="sub-title">Class Performance &amp; Results Report</div>
        </div>
        <div class="meta-bar">
            <div><strong>Class:</strong> {class_}{section} &nbsp;|&nbsp; <strong>Subject:</strong> {subject or 'All Subjects'}</div>
            <div><strong>Test No:</strong> {data.get('test_no') or 'All Tests'} &nbsp;|&nbsp; <strong>Date:</strong> {datetime.datetime.now().strftime('%d-%m-%Y')}</div>
        </div>
        <table>
            <thead>
                <tr><th>#</th><th>Admission No</th><th>Name</th><th>Class</th><th>House</th><th>Tests</th><th>Score</th></tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        <div class="footer">Generated by RRB CBT | {get_site_config()['developer_credit']}</div>
        <script>
            window.addEventListener('load', function() {{
                if (!window.location.search.includes('noprint')) {{
                    setTimeout(function() {{ window.print(); }}, 400);
                }}
            }});
        </script>
        </body></html>"""

        if export_format == 'download_pdf':
            pdf = HTML(string=html).write_pdf()
            resp = make_response(pdf)
            resp.headers['Content-Type'] = 'application/pdf'
            resp.headers['Content-Disposition'] = f'attachment; filename=class_report_{label}.pdf'
            return resp

        # Default: HTML Printable Preview
        return html

# ═══════════════════════════════════════════════════════════════
# BULLETIN BOARD
# ═══════════════════════════════════════════════════════════════

@app.route('/api/bulletins')
def api_bulletins():
    conn = get_db()
    c = conn.cursor()
    class_filter = request.args.get('class', '')
    query = """SELECT b.id, b.title, b.content, b.poster_type, b.target_class, b.created_at,
                      COALESCE(t.name,'Admin') as poster_name
               FROM bulletins b
               LEFT JOIN teachers t ON b.posted_by=t.id AND b.poster_type='teacher'
               WHERE b.is_active=1"""
    params = []
    if class_filter:
        query += " AND (b.target_class='' OR b.target_class=?)"
        params.append(class_filter)
    query += " ORDER BY b.created_at DESC LIMIT 15"
    c.execute(query, params)
    bulletins = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'bulletins': bulletins})

@app.route('/api/bulletins/post', methods=['POST'])
def post_bulletin():
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401
    title   = request.form.get('title','').strip()
    content = request.form.get('content','').strip()
    target  = request.form.get('target_class','').strip()
    if not title or not content:
        return jsonify({'status': 'error', 'message': 'Title and content required'}), 400
    if session.get('admin_logged_in'):
        posted_by   = 'admin'
        poster_type = 'admin'
    else:
        posted_by   = str(session.get('teacher_id'))
        poster_type = 'teacher'
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT INTO bulletins (title, content, posted_by, poster_type, target_class) VALUES (?,?,?,?,?)",
              (title, content, posted_by, poster_type, target))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Bulletin posted'})

@app.route('/api/bulletins/<int:bid>/delete', methods=['DELETE'])
def delete_bulletin(bid):
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error'}), 401
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE bulletins SET is_active=0 WHERE id=?", (bid,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

# ═══════════════════════════════════════════════════════════════
# ADMIN: UNLOCK STUDENT CLASS
# ═══════════════════════════════════════════════════════════════

@app.route('/admin/student/<student_id>/unlock_class', methods=['POST'])
@admin_required
def admin_unlock_student_class(student_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM student_class_lock WHERE student_id=?", (student_id,))
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'STUDENT_UNLOCK', 'student_class_lock', student_id, request.remote_addr)
    return jsonify({'status': 'success', 'message': 'Class lock removed. Student can re-register.'})

# ═══════════════════════════════════════════════════════════════
# ENHANCED AI TEST GENERATOR (Multi-type + Dual output mode)
# ═══════════════════════════════════════════════════════════════

@app.route('/admin/create_test', methods=['GET', 'POST'])
@admin_required
@limiter.limit("13 per minute")
def admin_create_test():
    """Feature #8: Admin test generation - unrestricted class/subject access."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, subject FROM questions WHERE class IS NOT NULL AND class!='' ORDER BY class, subject")
    all_pairs = [dict(r) for r in c.fetchall()]

    if request.method == 'POST':
        import urllib.request, json as json_lib

        board     = request.form.get('board', 'CBSE').strip() or 'CBSE'
        class_    = request.form.get('class','').strip()
        section   = request.form.get('section','').strip()
        subject   = request.form.get('subject','').strip()
        chapter   = request.form.get('chapter','').strip()
        test_no   = request.form.get('test_no','').strip()
        remark    = request.form.get('remark','').strip()
        output_mode = request.form.get('output_mode','cbt')
        method    = request.form.get('method','ai')

        if method == 'upload':
            eng_file = request.files.get('english_csv') or request.files.get('csv_file')
            hindi_file = request.files.get('hindi_csv')

            if not eng_file or not eng_file.filename.endswith('.csv'):
                conn.close()
                return jsonify({'status':'error','message':'English CSV file is required'}), 400

            # 1. Process compulsory English CSV & record inserted Row IDs
            eng_filename = secure_filename(eng_file.filename)
            eng_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"eng_{uuid.uuid4().hex[:8]}_{eng_filename}")
            eng_file.save(eng_filepath)

            inserted_ids = []
            try:
                with open(eng_filepath, 'r', encoding='utf-8-sig') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        parsed = parse_question_csv_row(row)
                        if not parsed or not parsed['question']:
                            continue
                        c.execute("""INSERT INTO questions
                                   (class, subject, chapter, test_no, question, option_a, option_b, option_c, option_d, correct_answer, board)
                                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                                 (class_, subject, chapter or test_no, test_no,
                                  parsed['question'], parsed['option_a'], parsed['option_b'],
                                  parsed['option_c'], parsed['option_d'], parsed['correct_answer'], board))
                        inserted_ids.append(c.lastrowid)
                conn.commit()
            except Exception as e:
                conn.close()
                return jsonify({'status':'error','message':f'Error parsing English CSV: {str(e)}'}), 400
            finally:
                if os.path.exists(eng_filepath):
                    os.remove(eng_filepath)

            if not inserted_ids:
                conn.close()
                return jsonify({'status':'error','message':'No valid questions found in English CSV'}), 400

            # 2. Process optional Hindi CSV & update matching English questions via Row index
            hindi_updated = 0
            if hindi_file and hindi_file.filename.endswith('.csv'):
                hin_filename = secure_filename(hindi_file.filename)
                hin_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"hin_{uuid.uuid4().hex[:8]}_{hin_filename}")
                hindi_file.save(hin_filepath)
                try:
                    with open(hin_filepath, 'r', encoding='utf-8-sig') as csvfile:
                        reader = csv.DictReader(csvfile)
                        for idx, row in enumerate(reader):
                            if idx < len(inserted_ids):
                                parsed = parse_question_csv_row(row)
                                if parsed and parsed['question']:
                                    c.execute("""UPDATE questions SET
                                               question_hi=?, option_a_hi=?, option_b_hi=?, option_c_hi=?, option_d_hi=?
                                               WHERE id=?""",
                                             (parsed['question'], parsed['option_a'], parsed['option_b'],
                                              parsed['option_c'], parsed['option_d'], inserted_ids[idx]))
                                    hindi_updated += 1
                    conn.commit()
                except Exception as e:
                    app.logger.warning(f"Error parsing Hindi CSV: {str(e)}")
                finally:
                    if os.path.exists(hin_filepath):
                        os.remove(hin_filepath)

            # 3. Record test generation history & test paper entries
            uploader_type = 'admin' if session.get('admin_logged_in') else 'teacher'
            uploader_id = 'admin' if uploader_type == 'admin' else str(session.get('teacher_id') or '0')
            teacher_name = uploader_type.capitalize()
            t_id = session.get('teacher_id')
            if t_id:
                c.execute("SELECT name FROM teachers WHERE id=?", (t_id,))
                t_r = c.fetchone()
                if t_r: teacher_name = t_r['name']

            c.execute("""INSERT INTO test_generation_history
                         (teacher_id,class,section,subject,chapter,test_no,output_mode,
                          total_questions,mcq_count,assertion_count,very_short_count,
                          short_count,long_count,case_study_count,remark,board)
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                      (uploader_id, class_, section, subject, chapter or test_no, test_no, output_mode,
                       len(inserted_ids), len(inserted_ids), 0, 0, 0, 0, 0, remark or 'CSV Upload', board))

            c.execute("""INSERT INTO mcq_test_history (teacher_id, teacher_name, class, section, subject, test_no, question_count)
                         VALUES (?,?,?,?,?,?,?)""",
                      (t_id or 0, teacher_name, class_, section, subject, test_no, len(inserted_ids)))

            if output_mode in ('cbt', 'both'):
                c.execute("""INSERT INTO test_papers
                             (filename,class,section,subject,test_no,uploaded_by,uploader_type,question_count,is_active,board)
                             VALUES (?,?,?,?,?,?,?,?,1,?)""",
                          (f"{class_}{section}_{subject}_{test_no}", class_, section, subject, test_no,
                           uploader_id, uploader_type, len(inserted_ids), board))

            conn.commit()
            conn.close()
            msg = f"{len(inserted_ids)} English questions uploaded"
            if hindi_updated > 0:
                msg += f" (with {hindi_updated} Hindi translations mapped)"
            return jsonify({'status':'success','message': msg})

        mcq_count       = int(request.form.get('mcq_count', 0) or 0)
        assertion_count = int(request.form.get('assertion_count', 0) or 0)

        type_counts = []
        if mcq_count:       type_counts.append(('MCQ', mcq_count))
        if assertion_count: type_counts.append(('Assertion-Reason', assertion_count))

        if not type_counts:
            conn.close()
            return jsonify({'status':'error','message':'Enter at least 1 question quantity'}), 400

        # Check duplicate test_no upfront to avoid waiting on AI generation
        if output_mode in ('cbt', 'both'):
            c.execute("SELECT COUNT(*) as cnt FROM questions WHERE class=? AND subject=? AND test_no=?",
                     (class_, subject, test_no))
            if c.fetchone()['cnt'] > 0:
                conn.close()
                return jsonify({'status':'error',
                                'message':f'Test No "{test_no}" already exists for {subject} Class {class_}. Use a different Test No.'}), 400

        # ── AI generation (batched to prevent token truncation) ──
        try:
            questions, gen_err = generate_cbt_questions_in_batches(
                class_, section, subject, chapter, test_no, remark,
                type_counts=type_counts, board=board, max_batch_size=10
            )
            if not questions:
                conn.close()
                return jsonify({'status':'error','message': gen_err or 'Failed to generate questions'}), 500

            c.execute("""INSERT INTO test_generation_history
                         (teacher_id,class,section,subject,chapter,test_no,output_mode,
                          total_questions,mcq_count,assertion_count,very_short_count,
                          short_count,long_count,case_study_count,remark,board)
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                      ('admin',class_,section,subject,chapter,test_no,output_mode,
                       len(questions),mcq_count,assertion_count,0,0,0,0,remark,board))
            c.execute("""INSERT INTO mcq_test_history (teacher_id, teacher_name, class, section, subject, test_no, question_count)
                         VALUES (0, 'Admin', ?, ?, ?, ?, ?)""",
                      (class_, section, subject, test_no, len(questions)))

            if output_mode in ('cbt', 'both'):
                inserted = 0
                for q in questions:
                    if q.get('question_type','MCQ') in ('MCQ','Assertion-Reason'):
                        img_path = process_cbt_question_image(q, class_, section, subject, chapter)

                        c.execute("""INSERT INTO questions
                                   (class,subject,chapter,test_no,question_type,question,question_hi,
                                    option_a,option_a_hi,option_b,option_b_hi,option_c,option_c_hi,
                                    option_d,option_d_hi,correct_answer,image_path,board)
                                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                 (class_,subject,chapter,test_no,q.get('question_type','MCQ'),
                                  q.get('question',''),q.get('question_hi',''),
                                  q.get('option_a',''),q.get('option_a_hi',''),
                                  q.get('option_b',''),q.get('option_b_hi',''),
                                  q.get('option_c',''),q.get('option_c_hi',''),
                                  q.get('option_d',''),q.get('option_d_hi',''),
                                  q.get('correct_answer',''),img_path,board))
                        inserted += 1
                c.execute("""INSERT INTO test_papers
                             (filename,class,section,subject,test_no,uploaded_by,uploader_type,question_count,is_active,board)
                             VALUES (?,?,?,?,?,?,?,?,1,?)""",
                          (f"{class_}{section}_{subject}_{test_no}",class_,section,subject,test_no,
                           'admin','admin',inserted,board))
                conn.commit()

            conn.commit()
            history_id = c.lastrowid
            conn.close()

            if output_mode == 'print':
                return jsonify({'status':'success','mode':'print',
                                'questions':questions,'history_id':history_id,
                                'message':f'{len(questions)} questions generated'})
            return jsonify({'status':'success','mode':'cbt',
                            'message':f'{len(questions)} questions generated. Test No: {test_no}'})

        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8') if hasattr(e, 'read') else str(e)
            app.logger.error(f"Gemini HTTP error: {err_body}")
            return jsonify({'status':'error','message':f'Gemini API error: {e.code} - {err_body[:200]}'}), 500
        except Exception as e:
            app.logger.error(f"Gemini exception: {str(e)}")
            return jsonify({'status':'error','message':str(e)}), 500

    conn.close()
    category = request.args.get('category', '').strip()
    is_coding = (category == 'coding_test')
    return render_template('admin_create_test.html', all_pairs=all_pairs, category=category, is_coding=is_coding)

def teacher_has_access(teacher_id, class_name, subject_name):
    """
    Check if a teacher is allotted to the given class and subject.
    Admins are always granted access.
    """
    if has_request_context() and session.get('admin_logged_in'):
        return True
    if not teacher_id:
        return False
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute("""SELECT id FROM teacher_assignments 
                     WHERE teacher_id=? AND class=? AND LOWER(subject)=LOWER(?) LIMIT 1""",
                  (teacher_id, str(class_name).strip(), str(subject_name).strip()))
        row = c.fetchone()
        conn.close()
        return bool(row)
    except Exception as e:
        app.logger.error(f"Error checking teacher access: {e}")
        return True

@app.route('/teacher/print_test/<int:result_id>')
@teacher_required
def teacher_print_test(result_id):
    conn = get_db()
    c = conn.cursor()
    # Get the test result
    c.execute("SELECT * FROM results WHERE id=?", (result_id,))
    result = c.fetchone()
    if not result:
        conn.close()
        return "Test result not found", 404
    result = dict(result)

    # Check teacher access (optional, if you want to restrict)
    teacher_id = session['teacher_id']
    if not teacher_has_access(teacher_id, result['class'], result['subject']):
        conn.close()
        return "Access denied", 403

    # Fetch questions for this test's class/subject
    c.execute("SELECT id, question, option_a, option_b, option_c, option_d, correct_answer, image_path FROM questions WHERE class=? AND subject=? ORDER BY id",
              (result['class'], result['subject']))
    questions = [dict(row) for row in c.fetchall()]
    conn.close()

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path = get_setting('logo_path', '')
    logo_base64 = None
    if logo_path:
        abs_path = os.path.join(app.static_folder, logo_path)
        if os.path.exists(abs_path):
            with open(abs_path, "rb") as f:
                logo_bytes = f.read()
                ext = os.path.splitext(abs_path)[1].lower()
                mime = "image/png" if ext == ".png" else "image/jpeg"
                logo_base64 = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode('utf-8')}"

    total_marks = len(questions)
    return render_template('question_paper_print.html',
                           school_name=school_name,
                           logo_base64=logo_base64,
                           class_name=result['class'],
                           subject=result['subject'],
                           questions=questions,
                           total_marks=total_marks,
                           date=datetime.datetime.now().strftime('%d/%m/%Y'))

@app.route('/teacher/test_history')
@teacher_required
def teacher_test_history():
    """
    Render the page showing test history for teacher's classes.
    """
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT * FROM test_generation_history
                 WHERE teacher_id=? ORDER BY created_at DESC""", (str(teacher_id),))
    history = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('teacher_test_history.html', history=history)

@app.route('/api/test_generation_history/<int:hid>/pdf')
def test_gen_history_pdf(hid):
    if not session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        return redirect(url_for('teacher_login'))
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM test_generation_history WHERE id=?", (hid,))
    hist = c.fetchone()
    if not hist:
        conn.close()
        return "Not found", 404
    hist = dict(hist)
    c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                        correct_answer, image_path,
                        CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                 FROM questions
                 WHERE (class=? OR class=? OR class LIKE ?)
                   AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                   AND (test_no=? OR chapter=? OR test_no LIKE ? OR chapter LIKE ?)
                 ORDER BY id""",
             (hist['class'],
              hist['class'].replace('th','').replace('st','').replace('nd','').replace('rd',''),
              f"%{hist['class']}%",
              hist['subject'], f"%{hist['subject']}%",
              hist['test_no'], hist['test_no'],
              f"%{hist['test_no']}%", f"%{hist['test_no']}%"))
    questions = [dict(r) for r in c.fetchall()]

    if not questions:
        c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                            correct_answer, image_path,
                            CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                     FROM questions
                     WHERE (LOWER(class)=LOWER(?) OR class LIKE ?)
                       AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                     ORDER BY id""",
                 (hist['class'], f"%{hist['class']}%", hist['subject'], f"%{hist['subject']}%"))
        questions = [dict(r) for r in c.fetchall()]

    conn.close()

    for q in questions:
        q_type = str(q.get('question_type') or '').strip()
        if not q_type or q_type.upper() in ['MCQ', 'CHOICE', 'OBJECTIVE']:
            q['question_type'] = 'MCQ'
    school_name = get_setting('school_name','RRB Group of Schools')
    logo_path   = get_setting('logo_path','')
    logo_base64 = None
    if logo_path:
        abs_path = os.path.join(app.static_folder, logo_path)
        if os.path.exists(abs_path):
            with open(abs_path,'rb') as f:
                logo_bytes = f.read()
                ext = os.path.splitext(abs_path)[1].lower()
                mime = 'image/png' if ext=='.png' else 'image/jpeg'
                logo_base64 = f"data:{mime};base64,{base64.b64encode(logo_bytes).decode()}"
    rendered = render_template('question_paper_template_v2.html',
        school_name=school_name, logo_base64=logo_base64,
        class_name=f"{hist['class']}{hist['section']}", subject=hist['subject'],
        chapter=hist['chapter'], test_no=hist['test_no'],
        questions=questions, total_marks=len(questions),
        date=datetime.datetime.now().strftime('%d/%m/%Y'))
    pdf = HTML(string=rendered, base_url=request.base_url).write_pdf()
    resp = make_response(pdf)
    resp.headers['Content-Type'] = 'application/pdf'
    resp.headers['Content-Disposition'] = f'attachment; filename=test_{hist["test_no"]}.pdf'
    return resp

# ═══════════════════════════════════════════════════════════════
# ADMIN: Student class lock management
# ═══════════════════════════════════════════════════════════════

@app.route('/admin/student_locks')
@admin_required
def admin_student_locks():
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT l.student_id, l.class, l.section, l.locked_at, s.name
                 FROM student_class_lock l
                 LEFT JOIN students s ON l.student_id=s.student_id
                 ORDER BY l.class, l.section, s.name""")
    locks = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'locks': locks})

# ═══════════════════════════════════════════════════════════════
# MULTIMODAL: IMAGEN 3 IMAGE GENERATION PROXY
# ═══════════════════════════════════════════════════════════════

@app.route('/api/generate_diagram', methods=['POST'])
def generate_diagram():
    """Proxy endpoint: receives an image prompt, calls Gemini Imagen, returns base64 image."""
    if not session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401

    import urllib.request, json as json_lib

    data = request.get_json() or {}
    image_prompt = data.get('prompt', '').strip()
    if not image_prompt:
        return jsonify({'status': 'error', 'message': 'No prompt provided'}), 400

    gemini_api_key = _get_gemini_key()
    if not gemini_api_key:
        return jsonify({'status': 'error', 'message': 'GEMINI_API_KEY not set'}), 500

    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key={gemini_api_key}"
        payload = json_lib.dumps({
            "instances": [{"prompt": image_prompt}],
            "parameters": {"sampleCount": 1, "aspectRatio": "1:1", "safetyFilterLevel": "block_some"}
        }).encode('utf-8')
        req = urllib.request.Request(url, data=payload,
                                     headers={'Content-Type': 'application/json'}, method='POST')
        with urllib.request.urlopen(req, timeout=45) as resp:
            result = json_lib.loads(resp.read().decode('utf-8'))
        img_b64 = result['predictions'][0]['bytesBase64Encoded']
        mime = result['predictions'][0].get('mimeType', 'image/png')
        return jsonify({'status': 'success', 'image': f'data:{mime};base64,{img_b64}',
                        'prompt': image_prompt})
    except urllib.error.HTTPError as e:
        err = e.read().decode('utf-8')
        return jsonify({'status': 'error', 'message': f'Imagen API error: {err}'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

def _get_gemini_key():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'apikey.env')
    if os.path.exists(env_path):
        try:
            content = open(env_path, 'r', encoding='utf-8').read().strip()
            if content:
                if 'GEMINI_API_KEY=' in content:
                    key = content.split('GEMINI_API_KEY=', 1)[1].split('\n', 1)[0].strip().strip('"\'')
                else:
                    key = content.strip().strip('"\'')
                if key:
                    os.environ['GEMINI_API_KEY'] = key
                    return key
        except Exception:
            pass
    return os.environ.get('GEMINI_API_KEY', '').strip().strip('"\'')

# ═══════════════════════════════════════════════════════════════
# BULLETIN POSTING API for Admin & Teacher dashboards
# ═══════════════════════════════════════════════════════════════

@app.route('/admin/bulletins')
@admin_required
def admin_bulletins_page():
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT b.*, COALESCE(t.name,'Admin') as poster_name
                 FROM bulletins b LEFT JOIN teachers t ON b.posted_by=t.id AND b.poster_type='teacher'
                 WHERE b.is_active=1 ORDER BY b.created_at DESC""")
    bulletins = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('admin_bulletins.html', bulletins=bulletins)


@app.route('/admin/audit_logs', methods=['GET'])
@admin_required
def admin_audit_logs():
    """
    Render Admin Audit Logs monitoring dashboard with filters.
    Supports filtering by user_type, action, start_date, and end_date.
    """
    user_type = request.args.get('user_type', '').strip().lower()
    action = request.args.get('action', '').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    
    conn = get_db()
    c = conn.cursor()
    
    query = "SELECT * FROM audit_logs WHERE 1=1"
    params = []
    
    if user_type:
        query += " AND LOWER(user_type) = ?"
        params.append(user_type)
        
    if action:
        query += " AND action = ?"
        params.append(action)
        
    if start_date:
        query += " AND timestamp >= ?"
        params.append(f"{start_date} 00:00:00")
        
    if end_date:
        query += " AND timestamp <= ?"
        params.append(f"{end_date} 23:59:59")
        
    query += " ORDER BY timestamp DESC LIMIT 500"
    
    c.execute(query, params)
    logs = [dict(r) for r in c.fetchall()]
    
    c.execute("SELECT COUNT(*) FROM audit_logs")
    total_count = c.fetchone()[0] or 0
    
    c.execute("SELECT COUNT(*) FROM audit_logs WHERE LOWER(user_type)='admin'")
    admin_count = c.fetchone()[0] or 0
    
    c.execute("SELECT COUNT(*) FROM audit_logs WHERE LOWER(user_type)='teacher'")
    teacher_count = c.fetchone()[0] or 0
    
    c.execute("SELECT COUNT(*) FROM audit_logs WHERE LOWER(user_type)='student'")
    student_count = c.fetchone()[0] or 0
    
    conn.close()
    
    if request.args.get('format') == 'json':
        return jsonify({'status': 'success', 'logs': logs, 'total': total_count})
        
    return render_template(
        'admin_audit_logs.html',
        logs=logs,
        total_count=total_count,
        admin_count=admin_count,
        teacher_count=teacher_count,
        student_count=student_count,
        selected_user_type=user_type,
        selected_action=action,
        start_date=start_date,
        end_date=end_date
    )


@app.route('/api/test_generation_history/<int:hid>/delete', methods=['DELETE'])
def delete_test_history(hid):
    """Feature #1: Delete individual test history entry."""
    if not session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authenticated'}), 401
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM test_generation_history WHERE id=?", (hid,))
    conn.commit()
    conn.close()
    u_type = 'teacher' if session.get('teacher_logged_in') else 'admin'
    u_id = session.get('teacher_id') if session.get('teacher_logged_in') else 'admin'
    log_audit_event(u_type, u_id, 'PAPER_DELETE', 'test_generation_history', hid, request.remote_addr)
    return jsonify({'status': 'success'})

@app.route('/api/test_generation_history/<int:hid>/print_preview')
def test_gen_print_preview(hid):
    """Feature #6: Return HTML page for browser print dialog (client-side PDF)."""
    if not session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        return redirect(url_for('teacher_login'))
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT * FROM test_generation_history WHERE id=?", (hid,))
    hist = c.fetchone()
    if not hist:
        conn.close()
        return "Not found", 404
    hist = dict(hist)

    # Get questions by test_no OR chapter using multi-tier fallback
    clean_cls = hist['class'].replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                        correct_answer, image_path,
                        CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                 FROM questions
                 WHERE (class=? OR class=? OR class LIKE ?)
                   AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                   AND (test_no=? OR chapter=? OR test_no LIKE ? OR chapter LIKE ?)
                 ORDER BY id""",
             (hist['class'], clean_cls, f"%{clean_cls}%",
              hist['subject'], f"%{hist['subject']}%",
              hist['test_no'], hist.get('chapter',''),
              f"%{hist['test_no']}%", f"%{hist.get('chapter','')}%"))
    questions = [dict(r) for r in c.fetchall()]

    if not questions:
        c.execute("""SELECT id, question, option_a, option_b, option_c, option_d,
                            correct_answer, image_path,
                            CASE WHEN question_type IS NULL OR trim(question_type) = '' THEN 'MCQ' ELSE question_type END as question_type
                     FROM questions
                     WHERE (LOWER(class)=LOWER(?) OR class LIKE ?)
                       AND (LOWER(subject)=LOWER(?) OR subject LIKE ?)
                     ORDER BY id""",
                 (hist['class'], f"%{hist['class']}%", hist['subject'], f"%{hist['subject']}%"))
        questions = [dict(r) for r in c.fetchall()]
    conn.close()

    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path   = get_setting('logo_path', '')
    logo_base64 = None
    if logo_path:
        abs_path = os.path.join(app.static_folder, logo_path)
        if os.path.exists(abs_path):
            with open(abs_path, 'rb') as f:
                lb = f.read()
                ext  = os.path.splitext(abs_path)[1].lower()
                mime = 'image/png' if ext == '.png' else 'image/jpeg'
                logo_base64 = f"data:{mime};base64,{base64.b64encode(lb).decode()}"

    return render_template('print_preview.html',
                           school_name=school_name,
                           logo_base64=logo_base64,
                           hist=hist,
                           questions=questions,
                           date=datetime.datetime.now().strftime('%d/%m/%Y'))

# ── TEACHER TEST GENERATION (v2) ───────────────────────────────────────

@app.route('/teacher/create_test_v2', methods=['GET', 'POST'])
@teacher_required
@limiter.limit("13 per minute")
def teacher_create_test_v2():
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT DISTINCT class, section, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
    assignments = [dict(r) for r in c.fetchall()]

    if request.method == 'POST':
        import urllib.request, json as json_lib

        board    = request.form.get('board', 'CBSE').strip() or 'CBSE'
        class_   = request.form.get('class','').strip()
        section  = request.form.get('section','').strip()
        subject  = request.form.get('subject','').strip()
        chapter  = request.form.get('chapter','').strip()
        test_no  = request.form.get('test_no','').strip()
        remark   = request.form.get('remark','').strip()
        output_mode = request.form.get('output_mode','cbt')
        method   = request.form.get('method','ai')

        mcq_count        = int(request.form.get('mcq_count', 0) or 0)
        assertion_count  = int(request.form.get('assertion_count', 0) or 0)
        very_short_count = int(request.form.get('very_short_count', 0) or 0)
        short_count      = int(request.form.get('short_count', 0) or 0)
        long_count       = int(request.form.get('long_count', 0) or 0)
        case_study_count = int(request.form.get('case_study_count', 0) or 0)

        if method == 'upload':
            eng_file = request.files.get('english_csv') or request.files.get('csv_file')
            hindi_file = request.files.get('hindi_csv')

            if not eng_file or not eng_file.filename.endswith('.csv'):
                conn.close()
                return jsonify({'status':'error','message':'English CSV file is required'}), 400

            eng_filename = secure_filename(eng_file.filename)
            eng_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"eng_{uuid.uuid4().hex[:8]}_{eng_filename}")
            eng_file.save(eng_filepath)

            inserted_ids = []
            try:
                with open(eng_filepath, 'r', encoding='utf-8-sig') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        parsed = parse_question_csv_row(row)
                        if not parsed or not parsed['question']:
                            continue
                        c.execute("""INSERT INTO questions
                                   (class,subject,chapter,test_no,question_type,question,
                                    option_a,option_b,option_c,option_d,correct_answer,board)
                                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                                 (class_, subject, chapter, test_no, parsed.get('question_type','MCQ'), parsed['question'],
                                  parsed['option_a'], parsed['option_b'], parsed['option_c'], parsed['option_d'],
                                  parsed['correct_answer'], board))
                        inserted_ids.append(c.lastrowid)
                conn.commit()
            finally:
                if os.path.exists(eng_filepath):
                    os.remove(eng_filepath)

            if not inserted_ids:
                conn.close()
                return jsonify({'status':'error','message':'No valid questions found in English CSV'}), 400

            hindi_updated = 0
            if hindi_file and hindi_file.filename and hindi_file.filename.endswith('.csv'):
                hin_filename = secure_filename(hindi_file.filename)
                hin_filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"hin_{uuid.uuid4().hex[:8]}_{hin_filename}")
                hindi_file.save(hin_filepath)
                try:
                    with open(hin_filepath, 'r', encoding='utf-8-sig') as csvfile:
                        reader = csv.DictReader(csvfile)
                        for idx, row in enumerate(reader):
                            if idx >= len(inserted_ids):
                                break
                            parsed_hi = parse_question_csv_row(row)
                            if not parsed_hi:
                                continue
                            q_id = inserted_ids[idx]
                            c.execute("""UPDATE questions
                                         SET question_hi=?, option_a_hi=?, option_b_hi=?, option_c_hi=?, option_d_hi=?
                                         WHERE id=?""",
                                     (parsed_hi['question'], parsed_hi['option_a'], parsed_hi['option_b'],
                                      parsed_hi['option_c'], parsed_hi['option_d'], q_id))
                            hindi_updated += 1
                    conn.commit()
                finally:
                    if os.path.exists(hin_filepath):
                        os.remove(hin_filepath)

            uploader_type = 'admin' if session.get('admin_logged_in') else 'teacher'
            uploader_id = 'admin' if uploader_type == 'admin' else str(session.get('teacher_id') or '0')
            teacher_name = uploader_type.capitalize()
            t_id = session.get('teacher_id')
            if t_id:
                c.execute("SELECT name FROM teachers WHERE id=?", (t_id,))
                t_r = c.fetchone()
                if t_r: teacher_name = t_r['name']

            c.execute("""INSERT INTO test_generation_history
                         (teacher_id,class,section,subject,chapter,test_no,output_mode,
                          total_questions,mcq_count,assertion_count,very_short_count,
                          short_count,long_count,case_study_count,remark,board)
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                      (uploader_id, class_, section, subject, chapter or test_no, test_no, output_mode,
                       len(inserted_ids), len(inserted_ids), 0, 0, 0, 0, 0, remark or 'CSV Upload', board))

            c.execute("""INSERT INTO mcq_test_history (teacher_id, teacher_name, class, section, subject, test_no, question_count)
                         VALUES (?,?,?,?,?,?,?)""",
                      (t_id or 0, teacher_name, class_, section, subject, test_no, len(inserted_ids)))

            if output_mode in ('cbt', 'both'):
                c.execute("""INSERT INTO test_papers
                             (filename,class,section,subject,test_no,uploaded_by,uploader_type,question_count,is_active,board)
                             VALUES (?,?,?,?,?,?,?,?,1,?)""",
                          (f"{class_}{section}_{subject}_{test_no}", class_, section, subject, test_no,
                           uploader_id, uploader_type, len(inserted_ids), board))

            conn.commit()
            conn.close()

            msg = f"{len(inserted_ids)} English questions uploaded"
            if hindi_updated > 0:
                msg += f" (with {hindi_updated} Hindi translations mapped)"
            return jsonify({'status':'success','message': msg})

        # ── AI generation ──
        type_counts = []
        if mcq_count:        type_counts.append(('MCQ (4 options)', mcq_count))
        if assertion_count:  type_counts.append(('Assertion-Reason', assertion_count))
        if very_short_count: type_counts.append(('Very Short Answer (1 Mark)', very_short_count))
        if short_count:      type_counts.append(('Short Answer (3 Marks)', short_count))
        if long_count:       type_counts.append(('Long Answer (5 Marks)', long_count))
        if case_study_count: type_counts.append(('Case Study (4 Marks)', case_study_count))

        if not type_counts:
            conn.close()
            return jsonify({'status':'error','message':'Enter at least 1 question quantity'}), 400

        # Check duplicate test_no upfront to avoid waiting on AI generation
        if output_mode in ('cbt', 'both'):
            c.execute("""SELECT COUNT(*) as cnt FROM questions
                         WHERE class=? AND subject=? AND test_no=?""",
                     (class_, subject, test_no))
            if c.fetchone()['cnt'] > 0:
                conn.close()
                return jsonify({
                    'status': 'error',
                    'message': f'Test No "{test_no}" already has questions for {subject} Class {class_}. Use a different Test No.'
                }), 400

        # ── AI generation (batched to prevent token truncation) ──
        try:
            questions, gen_err = generate_cbt_questions_in_batches(
                class_, section, subject, chapter, test_no, remark,
                type_counts=type_counts, board=board, max_batch_size=10
            )
            if not questions:
                conn.close()
                return jsonify({'status':'error','message': gen_err or 'Failed to generate questions'}), 500

            if output_mode in ('cbt', 'both'):
                inserted = 0
                for q in questions:
                    if q.get('question_type','MCQ') in ('MCQ','Assertion-Reason'):
                        img_path = process_cbt_question_image(q, class_, section, subject, chapter)

                        c.execute("""INSERT INTO questions
                                   (class,subject,chapter,test_no,question_type,question,question_hi,
                                    option_a,option_a_hi,option_b,option_b_hi,option_c,option_c_hi,
                                    option_d,option_d_hi,correct_answer,image_path,board)
                                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                 (class_,subject,chapter,test_no,
                                  q.get('question_type','MCQ'),q.get('question',''),q.get('question_hi',''),
                                  q.get('option_a',''),q.get('option_a_hi',''),
                                  q.get('option_b',''),q.get('option_b_hi',''),
                                  q.get('option_c',''),q.get('option_c_hi',''),
                                  q.get('option_d',''),q.get('option_d_hi',''),
                                  q.get('correct_answer',''),img_path,board))
                        inserted += 1
                c.execute("""INSERT INTO test_papers
                             (filename,class,section,subject,test_no,uploaded_by,uploader_type,question_count,is_active,board)
                             VALUES (?,?,?,?,?,?,?,?,1,?)""",
                          (f"{class_}{section}_{subject}_{test_no}",class_,section,subject,test_no,
                           str(teacher_id),'teacher',inserted,board))
                conn.commit()

            c.execute("""INSERT INTO test_generation_history
                         (teacher_id,class,section,subject,chapter,test_no,output_mode,
                          total_questions,mcq_count,assertion_count,very_short_count,
                          short_count,long_count,case_study_count,remark,board)
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                      (str(teacher_id),class_,section,subject,chapter,test_no,output_mode,
                       len(questions),mcq_count,assertion_count,very_short_count,
                       short_count,long_count,case_study_count,remark,board))
            teacher_name = session.get('teacher_name') or 'Teacher'
            c.execute("""INSERT INTO mcq_test_history (teacher_id, teacher_name, class, section, subject, test_no, question_count)
                         VALUES (?,?,?,?,?,?,?)""",
                      (teacher_id or 0, teacher_name, class_, section, subject, test_no, len(questions)))
            conn.commit()
            history_id = c.lastrowid
            conn.close()

            if output_mode == 'print':
                return jsonify({'status':'success','mode':'print',
                                'preview_url': f'/api/test_generation_history/{history_id}/print_preview',
                                'history_id': history_id,
                                'message':f'{len(questions)} questions generated'})
            else:
                return jsonify({'status':'success','mode':'cbt',
                                'message':f'{len(questions)} questions generated. Test No: {test_no}'})

        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8') if hasattr(e, 'read') else str(e)
            app.logger.error(f"Gemini HTTP error: {err_body}")
            return jsonify({'status':'error','message':f'Gemini API error: {e.code} - {err_body[:200]}'}), 500
        except Exception as e:
            app.logger.error(f"Teacher test generation exception: {str(e)}")
            return jsonify({'status':'error','message':str(e)}), 500

    conn.close()
    # Build class -> section -> subjects map
    class_section_subject_map = {}
    class_subject_map = {}
    for a in assignments:
        cls  = str(a['class']).strip() if a['class'] is not None else ''
        sec  = str(a.get('section') or '').strip().upper()
        subj = str(a['subject']).strip() if a['subject'] is not None else ''

        if not cls or not subj:
            continue

        if cls not in class_subject_map:
            class_subject_map[cls] = []
        if subj not in class_subject_map[cls]:
            class_subject_map[cls].append(subj)

        if cls not in class_section_subject_map:
            class_section_subject_map[cls] = {'all_subjects': [], 'sections': {}}

        if subj not in class_section_subject_map[cls]['all_subjects']:
            class_section_subject_map[cls]['all_subjects'].append(subj)

        if sec:
            if sec not in class_section_subject_map[cls]['sections']:
                class_section_subject_map[cls]['sections'][sec] = []
            if subj not in class_section_subject_map[cls]['sections'][sec]:
                class_section_subject_map[cls]['sections'][sec].append(subj)

    # Fetch recent test generation and upload history
    conn = get_db()
    c = conn.cursor()
    if session.get('admin_logged_in'):
        c.execute("SELECT * FROM test_generation_history ORDER BY created_at DESC LIMIT 20")
    else:
        c.execute("SELECT * FROM test_generation_history WHERE teacher_id=? OR teacher_id='admin' ORDER BY created_at DESC LIMIT 20", (str(teacher_id),))
    history = [dict(r) for r in c.fetchall()]
    conn.close()

    category = request.args.get('category', '').strip()
    is_coding = (category == 'coding_test')
    return render_template('teacher_create_test_v2.html',
                           assignments=assignments,
                           class_subject_map=class_subject_map,
                           class_section_subject_map=class_section_subject_map,
                           history=history,
                           category=category,
                           is_coding=is_coding)
@app.route('/admin/test_history')
@admin_required
def admin_test_history():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM test_generation_history ORDER BY created_at DESC")
    history = [dict(r) for r in c.fetchall()]
    conn.close()
    return render_template('teacher_test_history.html', history=history, is_admin=True)

# ── SHARED AI PROMPT BUILDER ──────────────────────────────────────────────────

def _build_ai_prompt(class_, section, subject, chapter, test_no, remark, type_instructions, board='CBSE'):
    board_str = (board or 'CBSE').strip()
    subj_lower = (subject or '').strip().lower()
    is_hindi = any(h in subj_lower for h in ['hindi', 'हिन्दी', 'हिंदी'])
    is_english = any(e in subj_lower for e in ['english', 'अंग्रेजी', 'अंग्रेज़ी'])

    if is_hindi:
        persona = f"You are an expert {board_str} Hindi Subject Teacher and question paper generator."
        lang_rule = """- STRICT HINDI LANGUAGE RULE (MANDATORY):
  * The subject is HINDI. Therefore, you MUST generate ALL questions, options, and explanations in pure HINDI language ONLY (Devanagari script, e.g. "निम्नलिखित में से कौन सा शब्द...", "(क) ...", "(ख) ...").
  * 'question' MUST be in Hindi. 'question_hi' must also contain the same Hindi question.
  * 'option_a', 'option_b', 'option_c', 'option_d' MUST be in Hindi only.
  * 'option_a_hi', 'option_b_hi', 'option_c_hi', 'option_d_hi' must contain the corresponding Hindi options.
  * DO NOT write questions or options in English under any circumstance."""
    elif is_english:
        persona = f"You are an expert {board_str} English Subject Teacher and question paper generator."
        lang_rule = """- STRICT ENGLISH LANGUAGE RULE (MANDATORY):
  * The subject is ENGLISH. Generate ALL questions and options strictly in ENGLISH language.
  * DO NOT convert, translate, or output questions/options in Hindi.
  * Set 'question_hi', 'option_a_hi', 'option_b_hi', 'option_c_hi', 'option_d_hi' to empty strings ""."""
    elif is_language_subject(subject):
        persona = f"You are an expert {board_str} {subject} Subject Teacher and question paper generator."
        lang_rule = f"""- STRICT {subject.upper()} LANGUAGE RULE (MANDATORY):
  * The subject is {subject} (a Language Subject). Generate ALL questions, options, and explanations strictly in {subject} language only.
  * DO NOT convert, translate, or output questions/options in any other language.
  * Set 'question_hi', 'option_a_hi', 'option_b_hi', 'option_c_hi', 'option_d_hi' to empty strings ""."""
    else:
        persona = f"You are an expert {board_str} {subject} question paper generator following the official {board_str} syllabus."
        lang_rule = f"""- STRICT LANGUAGE RULE (MANDATORY):
  * Since the subject is {subject} (NOT Hindi), generate all primary questions and options ('question', 'option_a', 'option_b', 'option_c', 'option_d') strictly in ENGLISH.
  * DO NOT convert or translate the primary question/options into Hindi.
  * You may provide standard secondary translations strictly in the '_hi' fields ('question_hi', 'option_a_hi', 'option_b_hi', 'option_c_hi', 'option_d_hi') for student reference, but the primary question and option fields MUST remain in English."""

    return f"""{persona}
Board: {board_str} | Class: {class_}{section} | Subject: {subject} | Chapter/Topics: {chapter} | Test No: {test_no}
Teacher instructions: {remark if remark else 'Standard difficulty, balanced coverage'}

Generate EXACTLY these question types:
{chr(10).join(type_instructions)}

OUTPUT: Return ONLY a valid JSON array. No markdown. No text before or after.
Every object must have ALL keys:
question_type, content_type, question, question_hi, option_a, option_a_hi, option_b, option_b_hi, option_c, option_c_hi, option_d, option_d_hi, correct_answer, smiles, circuit_type, image_prompt, image_keywords, marks

RULES:
{lang_rule}
- MATH LaTeX RULES: Inline $...$, Block $$...$$
- Escape backslashes in JSON: \\\\frac not \\frac
- CHEMISTRY mhchem RULES: For chemical equations use \\ce{{}} inside $...$ (escape braces in JSON as {{}})
  e.g. "$\\\\ce{{H2SO4 + 2NaOH -> Na2SO4 + 2H2O}}$"
  e.g. "$\\\\ce{{CaCO3(s) -> CaO(s) + CO2(g)}}$" for state symbols
  e.g. "$\\\\ce{{N2 + 3H2 <=> 2NH3}}$" for equilibrium arrows
  e.g. "$\\\\ce{{Fe^{{2+}} + 2e- -> Fe}}$" for ionic equations
  For simple formulas use $\\\\ce{{H2O}}$, $\\\\ce{{H2SO4}}$, $\\\\ce{{Ca(OH)2}}$
  Set content_type="chemistry" for all chemistry questions.
- DIAGRAM / CIRCUIT / MOLECULE RULES:
  * For electrical circuit questions (resistors, battery, switches, logic gates), provide 'circuit_type' (e.g. 'series_resistors', 'parallel_resistors', 'series_parallel', 'wheatstone', 'logic_gates') and descriptive 'image_prompt'.
  * For chemical structures or organic chemistry, provide 'smiles' (e.g. benzene="c1ccccc1", ethanol="CCO", aspirin="CC(=O)Oc1ccccc1C(=O)O") or compound name.
  * For biology, physics, or general science diagrams, set 'image_prompt' and short search terms in 'image_keywords' (e.g. ["human eye", "refraction of light", "plant cell", "electric motor"]).
- NON-MCQ: set option_a/b/c/d="", option_a_hi/b_hi/c_hi/d_hi="" and correct_answer="N/A"
- Leave smiles="", circuit_type="", image_prompt="", image_keywords=[] when no diagram is needed.

Return ONLY the JSON array starting with [ and ending with ]."""


def generate_cbt_questions_in_batches(class_, section, subject, chapter, test_no, remark, type_counts, board='CBSE', max_batch_size=10):
    """
    Generates CBT questions in structured batches to avoid LLM token limits (e.g. Gemini 8192 / DeepSeek 4096).
    Guarantees full question counts (e.g. 50 MCQs) without response truncation.
    """
    total_requested = sum(cnt for _, cnt in type_counts if cnt > 0)
    if total_requested == 0:
        return [], "No questions requested"

    # If small request (<= 12), single batch is sufficient and fastest
    if total_requested <= 12:
        batches = [[(label, cnt) for label, cnt in type_counts if cnt > 0]]
    else:
        batches = []
        for label, cnt in type_counts:
            if cnt <= 0:
                continue
            rem = cnt
            while rem > 0:
                chunk = min(rem, max_batch_size)
                batches.append([(label, chunk)])
                rem -= chunk

    all_questions = []
    seen_stems = []
    total_batches = len(batches)
    last_batch_err = ""
    app.logger.info(f"Generating {total_requested} questions across {total_batches} AI batches (class={class_}{section}, subject={subject}, chapter={chapter})")

    for b_idx, batch_specs in enumerate(batches, start=1):
        type_instructions = []
        batch_q_count = 0
        batch_label = ""
        for label, cnt in batch_specs:
            batch_q_count += cnt
            batch_label = label
            if label.startswith(f"{cnt} "):
                type_instructions.append(label)
            else:
                type_instructions.append(f"{cnt} {label}")

        batch_context = f"Batch {b_idx} of {total_batches}."
        if total_batches > 1:
            batch_context += f" Generate exactly {batch_q_count} unique questions."
            if seen_stems:
                exclusions = "\n".join(f"- {s}" for s in seen_stems[-15:])
                batch_context += f" DO NOT repeat or duplicate any of these previously generated questions:\n{exclusions}"
            batch_context += " Ensure diverse, balanced syllabus coverage across all subtopics of the chapter."

        full_remark = f"{remark} | {batch_context}".strip(' |') if remark else batch_context
        ai_prompt = _build_ai_prompt(
            class_, section, subject, chapter, test_no,
            remark=full_remark,
            type_instructions=type_instructions,
            board=board
        )

        parsed_batch = None
        for attempt in range(3):
            try:
                result, provider, err_msg = generate_ai_content(ai_prompt, timeout=75)
                if not result:
                    last_batch_err = err_msg or "AI returned empty response"
                    time.sleep(2 * (attempt + 1))
                    continue

                raw_text = result['candidates'][0]['content']['parts'][0]['text'].strip()
                if raw_text.startswith('```'):
                    first_line_end = raw_text.find('\n')
                    if first_line_end != -1:
                        raw_text = raw_text[first_line_end+1:].strip()
                    else:
                        raw_text = raw_text.lstrip('`').strip()
                        if raw_text.startswith('json'):
                            raw_text = raw_text[4:].strip()
                if raw_text.endswith('```'):
                    raw_text = raw_text[:-3].strip()

                questions_sub = safe_json_loads(raw_text)
                if questions_sub and isinstance(questions_sub, list) and len(questions_sub) > 0:
                    parsed_batch = questions_sub
                    break
                else:
                    last_batch_err = "Failed to parse JSON array from AI response"
            except Exception as e:
                last_batch_err = str(e)
            time.sleep(2 * (attempt + 1))

        if parsed_batch:
            for q in parsed_batch:
                if not isinstance(q, dict):
                    continue
                # Normalize question_type if missing
                if not q.get('question_type'):
                    q['question_type'] = 'Assertion-Reason' if 'assertion' in batch_label.lower() else 'MCQ'
                # Normalize correct_answer
                raw_ans = q.get('correct_answer', '')
                letter = _normalize_correct_answer(raw_ans, q.get('option_a', ''), q.get('option_b', ''), q.get('option_c', ''), q.get('option_d', ''))
                if letter in ('A', 'B', 'C', 'D'):
                    q['correct_answer'] = f"option_{letter.lower()}"
                q_stem = (q.get('question') or '').strip()
                if q_stem:
                    seen_stems.append(q_stem[:70])
                all_questions.append(q)
            app.logger.info(f"Batch {b_idx}/{total_batches} generated {len(parsed_batch)} questions (Total collected so far: {len(all_questions)})")
        else:
            app.logger.error(f"Batch {b_idx}/{total_batches} failed after 3 attempts: {last_batch_err}")

        # Sleep briefly between batches to respect rate limits
        if b_idx < total_batches:
            time.sleep(1.2)

    # Top-up deficit check if minor shortage occurred
    deficit = total_requested - len(all_questions)
    if 1 <= deficit <= 20 and all_questions:
        app.logger.info(f"Topping up {deficit} missing questions...")
        topup_instructions = [f"{deficit} MCQ (4 options)"]
        exclusions = "\n".join(f"- {s}" for s in seen_stems[-15:])
        topup_remark = f"{remark} | Top-up exactly {deficit} questions. DO NOT repeat:\n{exclusions}".strip(' |')
        topup_prompt = _build_ai_prompt(class_, section, subject, chapter, test_no, topup_remark, topup_instructions, board=board)
        for t_attempt in range(2):
            try:
                res, prov, err = generate_ai_content(topup_prompt, timeout=60)
                if res:
                    raw = res['candidates'][0]['content']['parts'][0]['text'].strip()
                    if raw.startswith('```'):
                        first_line_end = raw.find('\n')
                        if first_line_end != -1:
                            raw = raw[first_line_end+1:].strip()
                        else:
                            raw = raw.lstrip('`').strip()
                            if raw.startswith('json'):
                                raw = raw[4:].strip()
                    if raw.endswith('```'):
                        raw = raw[:-3].strip()
                    topup_qs = safe_json_loads(raw)
                    if topup_qs and isinstance(topup_qs, list):
                        for q in topup_qs:
                            if not isinstance(q, dict): continue
                            if not q.get('question_type'): q['question_type'] = 'MCQ'
                            letter = _normalize_correct_answer(q.get('correct_answer',''), q.get('option_a',''), q.get('option_b',''), q.get('option_c',''), q.get('option_d',''))
                            if letter in ('A', 'B', 'C', 'D'): q['correct_answer'] = f"option_{letter.lower()}"
                            all_questions.append(q)
                        if len(all_questions) >= total_requested:
                            break
            except Exception as e:
                app.logger.warning(f"Top-up call attempt {t_attempt+1} failed: {e}")
            time.sleep(2)

    if len(all_questions) > total_requested:
        all_questions = all_questions[:total_requested]

    if not all_questions:
        return None, f"AI Generation Failed: {last_batch_err or 'No questions produced by AI'}"

    return all_questions, None


# ── REDIS + RQ (REDIS QUEUE) & THREAD FALLBACK QUEUE MANAGER ──────────────────
import queue
import threading
import redis
from rq import Queue as RQ_Queue
from rq.job import Job as RQ_Job

paper_job_results = {}
paper_job_lock = threading.Lock()
ai_paper_thread_queue = queue.Queue()

REDIS_AVAILABLE = False
redis_conn = None
rq_paper_queue = None

def init_redis_queue():
    global REDIS_AVAILABLE, redis_conn, rq_paper_queue
    try:
        r_host = os.environ.get('REDIS_HOST', 'localhost')
        r_port = int(os.environ.get('REDIS_PORT', 6379))
        r_db = int(os.environ.get('REDIS_DB', 0))
        r_client = redis.Redis(host=r_host, port=r_port, db=r_db, socket_timeout=1.5)
        r_client.ping()
        redis_conn = r_client
        rq_paper_queue = RQ_Queue('paper_generation', connection=redis_conn)
        REDIS_AVAILABLE = True
        app.logger.info("[+] Redis connected! Using Redis + RQ for background paper processing.")
    except Exception as e:
        REDIS_AVAILABLE = False
        app.logger.info(f"[!] Redis unavailable ({e}). Reverting to built-in Thread Queue worker.")

init_redis_queue()


def process_ai_paper_job(job_data):
    """
    Core AI Paper Generation Task executed by Redis RQ Worker or Thread Queue.
    Generates paper via Gemini/multi-provider, formats .doc repository file,
    and inserts alert into teacher_notifications table.
    """
    job_id = job_data.get('job_id')
    prompt = job_data.get('prompt')
    teacher_id = job_data.get('teacher_id')
    teacher_name = job_data.get('teacher_name') or 'Admin'
    class_ = job_data.get('class_')
    section = job_data.get('section', 'A')
    subject = job_data.get('subject')
    exam_type = job_data.get('exam_type', 'Exam')
    duration = job_data.get('duration', '3 Hours')
    max_marks = job_data.get('max_marks', '100')
    school_name = job_data.get('school_name', 'RRB Group of Schools')
    
    with paper_job_lock:
        paper_job_results[job_id] = {
            'status': 'processing',
            'progress': 40,
            'message': 'AI engine generating questions and solution layout...'
        }
        
    try:
        result, provider, err_msg = generate_ai_content(prompt, timeout=90)
        if not result:
            with paper_job_lock:
                paper_job_results[job_id] = {'status': 'error', 'message': err_msg or 'AI generation failed'}
            return
            
        raw = result['candidates'][0]['content']['parts'][0]['text'].strip()
        if '```' in raw:
            for part in raw.split('```'):
                part = part.strip()
                if part.startswith('json'): part = part[4:].strip()
                if part.startswith('{'): raw = part; break
                
        paper_data = safe_json_loads(raw)
        if not paper_data:
            with paper_job_lock:
                paper_job_results[job_id] = {'status': 'error', 'message': 'Failed to parse generated paper JSON format. Please try again.'}
            return
        
        # Clean metadata from questions
        if isinstance(paper_data, dict) and 'sections' in paper_data:
            for sec in paper_data.get('sections', []):
                for q in sec.get('questions', []):
                    if isinstance(q.get('question'), str):
                        q['question'] = re.sub(r'\{"?meta(data)?"?:?.*\}', '', q['question'], flags=re.IGNORECASE).strip()

        # Save to repository: paper/Class_<N>/Section_<X>/
        abs_path, doc_filename, rel_link, *rest = save_paper_to_repository(
            paper_data, class_, section, subject, exam_type, teacher_name,
            meta={'school_name': school_name, 'duration': duration, 'max_marks': max_marks}
        )
        
        doc_link = f"/api/shared_papers/download/{rel_link}"
        
        with paper_job_lock:
            paper_job_results[job_id] = {
                'status': 'success',
                'progress': 100,
                'paper': paper_data,
                'doc_link': doc_link,
                'filename': doc_filename,
                'provider': provider
            }
            
        # Log notification into teacher_notifications table
        if teacher_id:
            try:
                conn = get_db()
                c = conn.cursor()
                c.execute("""INSERT INTO teacher_notifications (teacher_id, title, message, link)
                             VALUES (?, ?, ?, ?)""",
                          (teacher_id, "Descriptive Paper Ready!",
                           f"Your {subject} ({exam_type}) paper for Class {class_} {section} has been generated.",
                           doc_link))
                conn.commit()
                conn.close()
            except Exception as ex:
                app.logger.error(f"Error logging notification: {ex}")
                
    except Exception as e:
        with paper_job_lock:
            paper_job_results[job_id] = {'status': 'error', 'message': str(e)}


def enqueue_ai_paper_job(job_data):
    """
    Enqueues AI paper generation job.
    Uses Redis + RQ if connected; otherwise falls back to thread queue.
    """
    job_id = job_data['job_id']
    with paper_job_lock:
        paper_job_results[job_id] = {
            'status': 'queued',
            'progress': 15,
            'message': 'Queued in background processing pipeline...'
        }
        
    if REDIS_AVAILABLE and rq_paper_queue:
        try:
            rq_job = rq_paper_queue.enqueue(
                process_ai_paper_job,
                job_data,
                job_id=job_id,
                job_timeout=180
            )
            return job_id, 'rq'
        except Exception as e:
            app.logger.error(f"RQ enqueue failed ({e}). Reverting to thread queue.")
            
    ai_paper_thread_queue.put(job_data)
    return job_id, 'thread'


def ai_paper_thread_worker():
    """Fallback thread worker daemon."""
    while True:
        try:
            job_data = ai_paper_thread_queue.get()
            if job_data is None:
                break
            process_ai_paper_job(job_data)
        except Exception as e:
            app.logger.error(f"Error in thread worker: {e}")
        finally:
            ai_paper_thread_queue.task_done()

thread_daemon = threading.Thread(target=ai_paper_thread_worker, daemon=True)
thread_daemon.start()


@app.route('/api/paper_job/<job_id>', methods=['GET'])
def get_paper_job_status(job_id):
    """Poll job status for background AI paper processing."""
    with paper_job_lock:
        res = paper_job_results.get(job_id)
        
    if res:
        return jsonify({'status': 'success', 'job': res})
        
    # Check RQ Job if Redis is available
    if REDIS_AVAILABLE and redis_conn:
        try:
            rq_j = RQ_Job.fetch(job_id, connection=redis_conn)
            if rq_j:
                return jsonify({
                    'status': 'success',
                    'job': {
                        'status': rq_j.get_status(),
                        'message': f"Job status: {rq_j.get_status()}"
                    }
                })
        except Exception:
            pass
            
    return jsonify({'status': 'error', 'message': 'Job ID not found'}), 404


def latex_to_html_math(text):
    """
    Converts LaTeX math expressions into native HTML formatting (fractions, square roots,
    Greek symbols, exponents, subscripts) suitable for MS Word (.doc) document export.
    """
    if not text or not isinstance(text, str):
        return text

    # ── mhchem \ce{...} chemical equations → HTML ─────────────────────────────
    def _convert_ce(m):
        eq = m.group(1)
        # State symbols: (s) (l) (g) (aq) → subscript
        eq = re.sub(r'\(s\)', '<sub>(s)</sub>', eq)
        eq = re.sub(r'\(l\)', '<sub>(l)</sub>', eq)
        eq = re.sub(r'\(g\)', '<sub>(g)</sub>', eq)
        eq = re.sub(r'\(aq\)', '<sub>(aq)</sub>', eq)
        # Arrows
        eq = eq.replace('<=>', ' &#8652; ')   # ⇌ equilibrium
        eq = eq.replace('->', ' &rarr; ')      # →
        eq = eq.replace('<-', ' &larr; ')      # ←
        # Ion charges: Fe^{2+} → Fe<sup>2+</sup>
        eq = re.sub(r'\^\{([^}]+)\}', r'<sup>\1</sup>', eq)
        eq = re.sub(r'\^([+\-\d])', r'<sup>\1</sup>', eq)
        # Subscripts: numeric digits after letter groups
        eq = re.sub(r'([A-Za-z)\]]+)(\d+)', r'\1<sub>\2</sub>', eq)
        # Electrons: e- → e<sup>-</sup>
        eq = re.sub(r'\be-\b', 'e<sup>-</sup>', eq)
        return eq

    text = re.sub(r'\\ce\{((?:[^{}]|\{[^{}]*\})+)\}', _convert_ce, text)

    replacements = [
        # Fractions: \frac{a}{b} -> inline fraction block
        (r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'<span style="display:inline-block;vertical-align:-0.4em;text-align:center;margin:0 2px;"><span style="border-bottom:1px solid #000;display:block;padding:0 2px;font-size:0.9em;">\1</span><span style="display:block;padding:0 2px;font-size:0.9em;">\2</span></span>'),
        # Square root: \sqrt{x}
        (r'\\sqrt\{([^{}]+)\}', r'&radic;<span style="border-top:1px solid #000;padding-top:1px;margin-left:1px;">\1</span>'),
        # Math functions (render as text)
        (r'\\int', '&#8747;'),
        (r'\\sum', '&#8721;'),
        (r'\\prod', '&#8719;'),
        (r'\\lim', 'lim'),
        (r'\\sin', 'sin'),
        (r'\\cos', 'cos'),
        (r'\\tan', 'tan'),
        (r'\\cot', 'cot'),
        (r'\\sec', 'sec'),
        (r'\\cosec|\\csc', 'cosec'),
        (r'\\log', 'log'),
        (r'\\ln', 'ln'),
        (r'\\hat\{([^{}]+)\}', r'\1&#x0302;'),
        # Vectors & bold text: \vec{E}, \mathbf{B}
        (r'\\vec\{([^{}]+)\}', r'<b>\1&#8407;</b>'),
        (r'\\mathbf\{([^{}]+)\}', r'<b>\1</b>'),
        (r'\\text\{([^{}]+)\}', r'\1'),
        # Spacing
        (r'\\[,;]', ' '),
        (r'\\quad', '&nbsp;&nbsp;'),
        # Proportional to
        (r'\\propto', '&prop;'),
        # Greek capital letters
        (r'\\Phi', '&Phi;'),
        (r'\\Psi', '&Psi;'),
        (r'\\Sigma', '&Sigma;'),
        (r'\\Pi', '&Pi;'),
        (r'\\Omega', '&Omega;'),
        (r'\\Gamma', '&Gamma;'),
        (r'\\Theta', '&Theta;'),
        (r'\\Lambda', '&Lambda;'),
        (r'\\Delta', '&Delta;'),
        # Greek lowercase letters
        (r'\\alpha', '&alpha;'),
        (r'\\beta', '&beta;'),
        (r'\\gamma', '&gamma;'),
        (r'\\delta', '&delta;'),
        (r'\\varepsilon_0|\\epsilon_0', '&epsilon;<sub>0</sub>'),
        (r'\\varepsilon|\\epsilon', '&epsilon;'),
        (r'\\theta', '&theta;'),
        (r'\\lambda', '&lambda;'),
        (r'\\mu_0', '&mu;<sub>0</sub>'),
        (r'\\mu', '&mu;'),
        (r'\\pi', '&pi;'),
        (r'\\rho', '&rho;'),
        (r'\\sigma', '&sigma;'),
        (r'\\tau', '&tau;'),
        (r'\\chi_m|\\chi', '&chi;'),
        (r'\\phi', '&phi;'),
        (r'\\omega', '&omega;'),
        (r'\\infty', '&infin;'),
        # Operators & symbols
        (r'\\pm', '&plusmn;'),
        (r'\\times', '&times;'),
        (r'\\div', '&divide;'),
        (r'\\leq?', '&le;'),
        (r'\\geq?', '&ge;'),
        (r'\\neq', '&ne;'),
        (r'\\approx', '&approx;'),
        (r'\\degree|(?<=\^)\\circ', '&deg;'),
        (r'\\cdot', '&middot;'),
        (r'\\rightarrow|\\to', '&rarr;'),
        (r'\\leftarrow', '&larr;'),
        (r'\\leftrightarrow', '&harr;'),
        (r'\\Rightarrow', '&rArr;'),
    ]

    def replace_exp(m):
        return f"{m.group(1)}<sup>{m.group(2)}</sup>"

    def replace_sub(m):
        return f"{m.group(1)}<sub>{m.group(2)}</sub>"

    result = text
    for _ in range(3):
        for pattern, repl in replacements:
            result = re.sub(pattern, repl, result)

    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)\^\{([^}]+)\}', replace_exp, result)
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)\^([0-9a-zA-Z])', replace_exp, result)
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)_\{([^}]+)\}', replace_sub, result)
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)_([0-9a-zA-Z])', replace_sub, result)

    result = re.sub(r'\\\[|\\\]|\\\(|\\\)', '', result)
    result = re.sub(r'\$\$|\$', '', result)

    return result

def save_paper_to_repository(paper_data, class_, section, subject, exam_type, teacher_name, meta=None, paper_type=None):
    """
    Saves generated paper into structured folder hierarchy:
    paper/Class_<N>/Section_<X>/<class>_<section>_<subject>_<paper_type>_<TeacherName>.docx / .doc
    Generates circuits, checks image bank, and creates DOCX and HTML documents.
    """
    clean_cls = (class_ or 'General').strip().replace(' ', '_')
    clean_sec = (section or 'A').strip().replace(' ', '_')
    clean_sub = (subject or 'Subject').strip().replace(' ', '_')
    clean_exam = (exam_type or 'Exam').strip().replace(' ', '_')
    clean_teacher = (teacher_name or 'Teacher').strip().replace(' ', '_')
    
    meta_info = meta or {}
    paper_type = paper_type or meta_info.get('paper_type') or 'Descriptive Paper'
    clean_paper_type = paper_type.strip().replace(' ', '_')
    is_graded = meta_info.get('is_graded') if 'is_graded' in meta_info else (paper_type in ('Descriptive Paper', 'DPP'))

    cls_folder = f"Class_{clean_cls}"
    sec_folder = f"Section_{clean_sec}"
    
    rel_dir = os.path.join('paper', cls_folder, sec_folder)
    abs_dir = os.path.join(app.root_path, rel_dir)
    os.makedirs(abs_dir, exist_ok=True)

    # 1. Process question images (circuits with schemdraw, NCERT image bank, placeholders)
    try:
        image_handler.process_question_images(paper_data, output_dir=abs_dir)
    except Exception as e:
        app.logger.error(f"Error processing question images: {e}")
    
    filename = f"{clean_cls}_{clean_sec}_{clean_sub}_{clean_paper_type}_{clean_teacher}.doc"
    abs_file_path = os.path.join(abs_dir, filename)
    rel_link = f"Class_{clean_cls}/Section_{clean_sec}/{filename}"
    
    meta_full = {
        'school_name': meta_info.get('school_name') or get_setting('school_name', 'RRB Group of Schools'),
        'school_address': meta_info.get('school_address') or get_setting('school_address', ''),
        'academic_session': meta_info.get('academic_session') or get_setting('academic_session', ''),
        'exam_type': exam_type,
        'paper_type': paper_type,
        'is_graded': is_graded,
        'class': class_,
        'subject': subject,
        'topic': meta_info.get('topic') or meta_info.get('topics') or '',
        'full_syllabus': meta_info.get('full_syllabus', False),
        'duration': meta_info.get('duration') or '3 Hours',
        'max_marks': meta_info.get('max_marks') or ('30' if is_graded else ''),
        'teacher_name': clean_teacher.replace('_', ' '),
        'logo_path': meta_info.get('logo_path') or get_setting('logo_path', '')
    }

    # 2. Build native .docx using python-docx with formatting, circuits, and highlighted placeholders
    docx_filename = f"{clean_cls}_{clean_sec}_{clean_sub}_{clean_paper_type}_{clean_teacher}.docx"
    abs_docx_path = os.path.join(abs_dir, docx_filename)
    docx_rel_link = f"Class_{clean_cls}/Section_{clean_sec}/{docx_filename}"
    try:
        image_handler.create_paper_docx(paper_data, meta_full, abs_docx_path)
    except Exception as e:
        app.logger.error(f"Error creating docx paper: {e}")
        docx_rel_link = None

    school_name = meta_full['school_name']
    school_addr = meta_full['school_address']
    academic_sess = meta_full['academic_session']
    duration = meta_full['duration']
    max_marks = meta_full['max_marks']
    topic_str = str(meta_full['topic']).strip()
    is_full_syl = meta_full['full_syllabus'] in (True, 'yes', 'true', '1') or topic_str.lower() in ('full syllabus', 'complete syllabus', 'all chapters', 'whole syllabus', 'none', '')

    def _format_html_code_and_math(text_val):
        if not text_val:
            return ""
        s = str(text_val)
        # 1. Parse multi-line code blocks ```lang ... ```
        code_pat = re.compile(r'```(?:([a-zA-Z0-9_+-]+))?\s*\n?(.*?)\n?```', re.DOTALL)
        code_blocks = []
        def _stash_code(m):
            lang = (m.group(1) or '').strip().upper()
            lang_badge = f"<div style='background:#1e1b4b;color:#a5b4fc;font-size:7.5pt;font-weight:bold;padding:1px 8px;display:inline-block;border-radius:3px 3px 0 0;'>{lang}</div>" if lang else ""
            escaped_code = m.group(2).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            code_blocks.append(f"<div style='margin:6px 0;'>{lang_badge}<pre style='background:#f8fafc;border:1px solid #cbd5e1;border-left:4px solid #4f46e5;border-radius:4px;padding:8px 12px;font-family:\"Consolas\",\"Courier New\",monospace;font-size:9.5pt;line-height:1.4;white-space:pre-wrap;overflow-x:auto;margin:0;text-align:left;color:#0f172a;'><code>{escaped_code}</code></pre></div>")
            return f"__CODE_BLOCK_{len(code_blocks)-1}__"
            
        s_stashed = code_pat.sub(_stash_code, s)

        # 2. Parse inline code `...`
        inline_code_pat = re.compile(r'`([^`\n]+)`')
        inline_blocks = []
        def _stash_inline(m):
            escaped = m.group(1).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            inline_blocks.append(f"<code style='background:#f1f5f9;color:#1e1b4b;border:1px solid #cbd5e1;padding:1px 4px;border-radius:3px;font-family:\"Consolas\",\"Courier New\",monospace;font-size:9.5pt;font-weight:600;'>{escaped}</code>")
            return f"__INLINE_CODE_{len(inline_blocks)-1}__"

        s_stashed = inline_code_pat.sub(_stash_inline, s_stashed)

        # 3. Convert math & chemistry on remaining text
        s_math = latex_to_html_math(s_stashed)

        # 4. Restore code blocks & inline code
        for idx, cb in enumerate(code_blocks):
            s_math = s_math.replace(f"__CODE_BLOCK_{idx}__", cb)
        for idx, ib in enumerate(inline_blocks):
            s_math = s_math.replace(f"__INLINE_CODE_{idx}__", ib)
            
        return s_math

    sections_html = ""
    if isinstance(paper_data, dict) and 'sections' in paper_data:
        for sec in paper_data.get('sections', []):
            label = _format_html_code_and_math(sec.get('section_label', ''))
            title = _format_html_code_and_math(sec.get('section_title', ''))
            instruction = _format_html_code_and_math(sec.get('instruction', ''))
            sections_html += f"<h3 class='sec-banner' style='margin-top:16px;color:#1e1b4b;border-bottom:1px solid #ddd;padding-bottom:4px;page-break-after:avoid;'>{label}: {title}</h3>"
            if instruction:
                sections_html += f"<p style='font-style:italic;color:#555;margin-bottom:8px;page-break-after:avoid;'>Note: {instruction}</p>"
            
            for q in sec.get('questions', []):
                q_num = q.get('number', '')
                q_text = _format_html_code_and_math(q.get('question', ''))
                marks_badge = f"<span style='float:right;font-weight:bold;color:#475569;'>[{sec.get('marks_per_question', 1)}M]</span>" if is_graded else ""
                sections_html += f"<div class='q-item' style='margin-bottom:10px;page-break-inside:avoid;'><strong>Q{q_num}.</strong> {q_text} {marks_badge}</div>"
                
                # Render generated circuit / NCERT image or placeholder in HTML doc
                if q.get('image_base64'):
                    sections_html += f"<div style='text-align:center;margin:8px 0;'><img src='{q['image_base64']}' style='max-width:320px;max-height:220px;' alt='Diagram'></div>"
                elif q.get('image_placeholder'):
                    sections_html += f"<div style='background:#fef08a;color:#854d0e;padding:6px 12px;border-radius:4px;display:inline-block;font-weight:bold;margin:6px 0;'>[IMAGE NEEDED: {q['image_placeholder']}]</div>"
                elif q.get('needs_image'):
                    sections_html += f"<div style='background:#fef08a;color:#854d0e;padding:6px 12px;border-radius:4px;display:inline-block;font-weight:bold;margin:6px 0;'>[IMAGE NEEDED: Diagram for Question {q_num}]</div>"

                opts = q.get('options', [])
                if opts:
                    sections_html += "<ul style='list-style-type:none;padding-left:15px;margin-top:4px;'>"
                    for opt in opts:
                        opt_text = _format_html_code_and_math(opt)
                        sections_html += f"<li style='margin-bottom:2px;'>{opt_text}</li>"
                    sections_html += "</ul>"
                
                sub_qs = q.get('sub_questions', [])
                if sub_qs:
                    sections_html += "<ol style='padding-left:20px;margin-top:4px;'>"
                    for sq in sub_qs:
                        sq_raw = sq if isinstance(sq, str) else sq.get('question', '')
                        sq_text = _format_html_code_and_math(sq_raw)
                        sections_html += f"<li>{sq_text}</li>"
                    sections_html += "</ol>"

    topic_html_line = f"<div style='font-size:10pt;font-weight:bold;color:#475569;margin-top:2px;'>TOPIC / CHAPTER: {topic_str}</div>" if (not is_full_syl and topic_str) else ""

    exam_header_title = f"{paper_type.upper()}{' — ' + exam_type.upper() if exam_type and exam_type.lower() != paper_type.lower() else ''}{f' (Academic Session: {academic_sess})' if academic_sess else ''}"
    marks_header_td = f"<td align='right'>Maximum Marks (MM): {max_marks}</td>" if is_graded and max_marks else "<td></td>"

    doc_html = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'><title>{subject} {paper_type}</title>
<style>
body {{ font-family: Arial, 'Times New Roman', serif; font-size: 11pt; line-height: 1.45; margin: 20px; color: #000; }}
.header {{ text-align: center; margin-bottom: 16px; border-bottom: 2px solid #000; padding-bottom: 10px; }}
.p-school {{ font-size: 16pt; font-weight: bold; text-transform: uppercase; color: #1e1b4b; margin-bottom: 2px; }}
.p-addr {{ font-size: 9.5pt; color: #475569; margin-bottom: 4px; }}
.p-exam {{ font-size: 11pt; font-weight: bold; margin-bottom: 2px; }}
.p-cls {{ font-size: 11pt; font-weight: bold; color: #1e1b4b; }}
pre, code {{ font-family: 'Consolas', 'Courier New', monospace; }}
@media print {{ body {{ margin: 10mm; }} .q-item {{ page-break-inside: avoid; }} .sec-banner {{ page-break-after: avoid; }} }}
</style>
</head>
<body>
<div class='header'>
    <div class='p-school'>{school_name}</div>
    {f"<div class='p-addr'>{school_addr}</div>" if school_addr else ""}
    <div class='p-exam'>{exam_header_title}</div>
    <div class='p-cls'>CLASS: {class_} &nbsp;|&nbsp; SUBJECT: {subject.upper()}</div>
    {topic_html_line}
</div>
<table width='100%' style='margin-bottom:15px;font-weight:bold;font-size:10pt;border-top:1px solid #000;border-bottom:1px solid #000;padding:5px 0;'>
    <tr>
        <td>Time Allowed: {duration}</td>
        {marks_header_td}
    </tr>
    <tr>
        <td>Teacher: {clean_teacher.replace('_', ' ')}</td>
        <td align='right'>Date: {datetime.datetime.now().strftime('%d/%m/%Y')}</td>
    </tr>
</table>
{sections_html}
<div style='margin-top:30px;text-align:center;font-size:9pt;font-style:italic;color:#666;border-top:1px solid #ccc;padding-top:8px;'>
    *** End of {paper_type} ***<br>Generated by RRB CBT &nbsp;&middot;&nbsp; {get_site_config()['developer_credit']}
</div>
</body>
</html>"""

    try:
        with open(abs_file_path, 'w', encoding='utf-8') as f:
            f.write(doc_html)
    except Exception as e:
        app.logger.error(f"Error saving paper doc: {e}")
        
    return abs_file_path, filename, rel_link, docx_rel_link


@app.route('/descriptive_paper', methods=['GET', 'POST'])
@app.route('/question_bank', methods=['GET', 'POST'])
@limiter.limit("30 per minute")
def descriptive_paper():
    """Unified Paper & Question Bank generator — accessible to both admin and teacher."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return redirect(url_for('teacher_login'))

    # GET — show the form
    if request.method == 'GET':
        teacher_id = session.get('teacher_id')
        assignments = []
        all_pairs   = []
        if teacher_id:
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT DISTINCT class, subject FROM teacher_assignments WHERE teacher_id=?", (teacher_id,))
            assignments = [dict(r) for r in c.fetchall()]
            conn.close()
        elif session.get('admin_logged_in'):
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT DISTINCT class, subject FROM questions WHERE class IS NOT NULL AND class!='' ORDER BY class, subject")
            all_pairs = [dict(r) for r in c.fetchall()]
            conn.close()

        school_name      = get_setting('school_name', 'RRB Group of Schools')
        school_address   = get_setting('school_address', '')
        academic_session = get_setting('academic_session', '')
        logo_path        = get_setting('logo_path', '')

        default_paper_type = 'Question Bank' if request.path == '/question_bank' else 'Descriptive Paper'
        req_type = request.args.get('type')
        if req_type:
            default_paper_type = req_type

        return render_template('descriptive_paper.html',
                               assignments=assignments,
                               all_pairs=all_pairs,
                               school_name=school_name,
                               school_address=school_address,
                               academic_session=academic_session,
                               logo_path=logo_path,
                               is_admin=bool(session.get('admin_logged_in')),
                               default_paper_type=default_paper_type)

    # POST — generate via Gemini and return structured JSON for the preview
    import urllib.request, json as json_lib

    # ── Collect all form fields ──────────────────────────────────────────────
    board          = request.form.get('board', 'CBSE').strip() or 'CBSE'
    class_         = request.form.get('class', '').strip()
    subject        = request.form.get('subject', '').strip()
    topics         = request.form.get('topics', '').strip()
    full_syllabus  = request.form.get('full_syllabus') == 'yes'
    exam_type      = request.form.get('exam_type', 'Unit Test').strip()
    paper_type     = request.form.get('paper_type', 'Descriptive Paper').strip() or 'Descriptive Paper'
    is_graded      = paper_type in ('Descriptive Paper', 'DPP')
    duration       = request.form.get('duration', '3 Hours').strip()
    if is_graded:
        max_marks  = request.form.get('max_marks', '30').strip() or '30'
    else:
        max_marks  = ''
    academic_sess  = request.form.get('academic_session', '').strip() or get_setting('academic_session', '')
    school_nm      = request.form.get('school_name', '').strip() or get_setting('school_name', 'RRB Group of Schools')
    school_addr    = request.form.get('school_address', '').strip() or get_setting('school_address', '')
    remarks        = request.form.get('remarks', '').strip()

    # Question counts and marks per type
    mcq_count        = int(request.form.get('mcq_count', 0) or 0)
    mcq_marks        = int(request.form.get('mcq_marks', 1) or 1) if is_graded else 1
    fib_count        = int(request.form.get('fib_count', 0) or 0)
    fib_marks        = int(request.form.get('fib_marks', 1) or 1) if is_graded else 1
    tf_count         = int(request.form.get('tf_count', 0) or 0)
    tf_marks         = int(request.form.get('tf_marks', 1) or 1) if is_graded else 1
    ar_count         = int(request.form.get('ar_count', 0) or 0)
    ar_marks         = int(request.form.get('ar_marks', 1) or 1) if is_graded else 1
    vs_count         = int(request.form.get('vs_count', 0) or 0)
    vs_marks         = int(request.form.get('vs_marks', 2) or 2) if is_graded else 1
    sh_count         = int(request.form.get('sh_count', 0) or 0)
    sh_marks         = int(request.form.get('sh_marks', 3) or 3) if is_graded else 1
    lg_count         = int(request.form.get('lg_count', 0) or 0)
    lg_marks         = int(request.form.get('lg_marks', 5) or 5) if is_graded else 1
    cs_count         = int(request.form.get('cs_count', 0) or 0)
    cs_marks         = int(request.form.get('cs_marks', 4) or 4) if is_graded else 1

    total_q = mcq_count + fib_count + tf_count + ar_count + vs_count + sh_count + lg_count + cs_count

    if total_q == 0:
        return jsonify({'status': 'error', 'message': 'Please enter at least 1 question'}), 400

    max_allowed_q = 250 if paper_type in ('Question Bank', 'Worksheet') else 80
    if total_q > max_allowed_q:
        return jsonify({'status': 'error', 'message': f'Total questions ({total_q}) exceeds limit of {max_allowed_q} for {paper_type}.'}), 400

    # ── Option: Generate from Exemplar Bank ───────────────────────────────────
    source_mode = request.form.get('source_mode', 'ai').strip().lower()
    if source_mode in ('exemplar', 'exemplar_bank'):
        ex_paper = exemplar_extractor.generate_paper_from_exemplar(
            subject=subject,
            class_num=class_,
            chapter_name=None if full_syllabus else topics,
            mcq_count=mcq_count,
            short_count=sh_count + vs_count,
            long_count=lg_count + cs_count
        )
        if ex_paper and ex_paper.get('sections') and any(s.get('questions') for s in ex_paper['sections']):
            return jsonify({
                'status': 'success',
                'provider': 'NCERT Exemplar Bank',
                'data': ex_paper,
                'paper_data': ex_paper
            })
        else:
            return jsonify({
                'status': 'error',
                'message': f'No questions found in Exemplar Bank for {subject} Class {class_} ({topics or "All Chapters"}). Please upload/process an Exemplar PDF first, or choose AI Generation mode.'
            }), 404


    try:
        # ── Optional image upload ────────────────────────────────────────────────
        image_context = ''
        img_file = request.files.get('image_file')
        if img_file and img_file.filename:
            allowed_img = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
            ext = img_file.filename.rsplit('.', 1)[-1].lower()
            if ext in allowed_img:
                img_bytes   = img_file.read()
                img_b64     = base64.b64encode(img_bytes).decode()
                img_mime    = f'image/{ext}' if ext != 'jpg' else 'image/jpeg'
                image_context = f'\n\nAn image has been provided (base64 omitted here). Use it conceptually to inspire questions involving diagrams, graphs, or visual-based problems relevant to {subject}.'

        # ── Build structured section specifications ──────────────────────────────
        section_specs = []
        if mcq_count:
            section_specs.append({
                'type': 'mcq',
                'title': 'Multiple Choice Questions (MCQ)',
                'count': mcq_count,
                'marks': mcq_marks,
                'instruction': 'Choose the correct option from (a), (b), (c), or (d).',
                'detail': 'Each question MUST have an "options" array with exactly 4 distinct choices: ["(a) ...", "(b) ...", "(c) ...", "(d) ..."]. Set "sub_questions": [].'
            })
        if fib_count:
            section_specs.append({
                'type': 'fib',
                'title': 'Fill in the Blanks',
                'count': fib_count,
                'marks': fib_marks,
                'instruction': 'Fill in the blanks with appropriate terms.',
                'detail': 'Each question statement MUST contain a blank line "_______" to fill. Set "options": []. Set "sub_questions": [].'
            })
        if tf_count:
            section_specs.append({
                'type': 'tf',
                'title': 'True / False',
                'count': tf_count,
                'marks': tf_marks,
                'instruction': 'State whether the following statements are True or False.',
                'detail': 'Clear factual statements. Set "options": ["True", "False"]. Set "sub_questions": [].'
            })
        if ar_count:
            section_specs.append({
                'type': 'ar',
                'title': 'Assertion-Reason',
                'count': ar_count,
                'marks': ar_marks,
                'instruction': 'Read the Assertion (A) and Reason (R) statements and choose the correct option.',
                'detail': 'State Assertion (A) and Reason (R) clearly. Provide standard 4 AR options: ["(a) Both A and R are true and R is the correct explanation of A", "(b) Both A and R are true but R is not the correct explanation of A", "(c) A is true but R is false", "(d) A is false but R is true"].'
            })
        if vs_count:
            section_specs.append({
                'type': 'vs',
                'title': 'Very Short Answer',
                'count': vs_count,
                'marks': vs_marks,
                'instruction': 'Answer in one or two sentences.',
                'detail': 'Direct, concise conceptual questions. Set "options": []. Set "sub_questions": [].'
            })
        if sh_count:
            section_specs.append({
                'type': 'sh',
                'title': 'Short Answer',
                'count': sh_count,
                'marks': sh_marks,
                'instruction': 'Answer in 30 to 50 words.',
                'detail': 'Conceptual, numerical, or analytical questions. Set "options": []. Set "sub_questions": [].'
            })
        if lg_count:
            section_specs.append({
                'type': 'lg',
                'title': 'Long Answer',
                'count': lg_count,
                'marks': lg_marks,
                'instruction': 'Answer in detail with diagrams, steps, or derivations where applicable.',
                'detail': 'Comprehensive questions, derivations, or multi-part problems. Set "options": [].'
            })
        if cs_count:
            section_specs.append({
                'type': 'cs',
                'title': 'Case Study Based',
                'count': cs_count,
                'marks': cs_marks,
                'instruction': 'Read the case passage carefully and answer the following sub-questions.',
                'detail': 'In "question", provide a 3-5 sentence contextual passage/case scenario. In "sub_questions", provide an array of 3-4 sub-questions (e.g. [{"question": "(i) ..."}, {"question": "(ii) ..."}]).'
            })

        subj_lower = (subject or '').strip().lower()
        is_hindi = any(h in subj_lower for h in ['hindi', 'हिन्दी', 'हिंदी'])
        is_english = any(e in subj_lower for e in ['english', 'अंग्रेजी', 'अंग्रेज़ी'])
        is_math = any(m in subj_lower for m in ['math', 'maths', 'mathematics', 'statistics', 'calculus', 'algebra', 'geometry', 'trigonometry'])
        is_accounts = any(a in subj_lower for a in ['account', 'accountancy', 'bookkeeping', 'commerce', 'economics', 'business studies', 'finance'])
        is_chemistry = any(c in subj_lower for c in ['chemistry', 'biochemistry'])
        is_physics = any(p in subj_lower for p in ['physics', 'physical science'])
        is_programming = any(pr in subj_lower for pr in ['computer', 'java', 'python', 'programming', 'c++', 'sql', 'html', 'coding', 'information technology', 'it', 'computer science', 'computer application'])

        if is_hindi:
            lang_rule_paper = """- STRICT HINDI LANGUAGE DIRECTIVE (MANDATORY):
  * Since the subject is HINDI, the ENTIRE question paper (all questions, options, sub_questions, section instructions, and passages) MUST be generated in HINDI language ONLY (Devanagari script). DO NOT output in English."""
        elif is_english:
            lang_rule_paper = """- STRICT ENGLISH LANGUAGE DIRECTIVE (MANDATORY):
  * Since the subject is ENGLISH, the ENTIRE question paper MUST be generated in ENGLISH language only. DO NOT convert or translate into Hindi."""
        elif is_language_subject(subject):
            lang_rule_paper = f"""- STRICT {subject.upper()} LANGUAGE DIRECTIVE (MANDATORY):
  * Since the subject is {subject} (a Language Subject), the ENTIRE question paper (all questions, options, sub_questions, section instructions, and passages) MUST be generated in {subject} language only. DO NOT convert or translate into Hindi or English."""
        else:
            lang_rule_paper = f"""- STRICT LANGUAGE DIRECTIVE (MANDATORY):
  * Since the subject is {subject} (NOT Hindi), generate the entire question paper strictly in ENGLISH (or the standard subject language). DO NOT convert or translate questions into Hindi."""

        topic_text = 'Complete Syllabus' if full_syllabus else (topics or 'All Topics')
        grading_line = f"- Maximum Marks: {max_marks}" if is_graded and max_marks else "- Marks: Not Applicable (Practice material - ungraded)"

        # ── Group requested sections into balanced batches (max ~20-25 questions per batch) ───
        # This prevents LLM output token cutoff and guarantees ALL sections (A through H) are fully generated.
        batches = []
        current_batch = []
        current_q_count = 0

        for spec in section_specs:
            if spec['count'] > 25:
                if current_batch:
                    batches.append(current_batch)
                    current_batch = []
                    current_q_count = 0
                
                remaining = spec['count']
                chunk_num = 1
                while remaining > 0:
                    take = min(remaining, 20)
                    sub_spec = dict(spec)
                    sub_spec['count'] = take
                    if spec['count'] > 25:
                        sub_spec['title'] = f"{spec['title']} (Part {chunk_num})"
                    batches.append([sub_spec])
                    remaining -= take
                    chunk_num += 1
            else:
                if current_q_count + spec['count'] > 25 and current_batch:
                    batches.append(current_batch)
                    current_batch = [spec]
                    current_q_count = spec['count']
                else:
                    current_batch.append(spec)
                    current_q_count += spec['count']

        if current_batch:
            batches.append(current_batch)

        import math as _math

        def generate_single_batch(batch_specs, b_idx):
            lines = []
            for s in batch_specs:
                m_str = f", {s['marks']} mark{'s' if s['marks'] > 1 else ''} each" if is_graded else ""
                orig_count = s['count']
                # Generate 25% extra questions so the paper has choice (attempt any N out of N+extra)
                # Only apply for non-MCQ/TF/FIB where choice makes sense (short, long, case study, VS, AR)
                choice_types = ('vs', 'sh', 'lg', 'cs', 'ar')
                if s.get('type') in choice_types and orig_count >= 2:
                    extra_count = _math.ceil(orig_count * 1.25)
                    choice_note = f" (Generate EXACTLY {extra_count} questions; students will attempt any {orig_count} out of {extra_count})"
                else:
                    extra_count = orig_count
                    choice_note = ""
                lines.append(f"• Section: {s['title']} — EXACTLY {extra_count} questions{m_str}.{choice_note}\n  Instruction: {s['instruction']}{(' Attempt any ' + str(orig_count) + ' out of ' + str(extra_count) + '.') if extra_count > orig_count else ''}\n  Requirements: {s['detail']}")

            batch_prompt = f"""You are an expert {board} question paper setter.
Generate ONLY the following specific section(s) of a {paper_type.lower()} strictly adhering to the LATEST 2024-25 / 2025-26 rationalized curriculum.

PAPER DETAILS:
- Paper Type: {paper_type}
- Board: {board}
- Class: {class_}
- Subject: {subject}
- Topics/Chapters: {topic_text}
{grading_line}

SECTIONS TO GENERATE IN THIS CALL (MANDATORY: GENERATE EVERY QUESTION LISTED BELOW):
{chr(10).join(lines)}

TEACHER INSTRUCTIONS / REMARKS:
{remarks if remarks else f'Standard board difficulty. High quality questions. Follow latest {board} pattern.'}
{image_context}

SUBJECT-SPECIFIC FORMATTING DIRECTIVES (STRICT COMPLIANCE — apply whichever is relevant):
1. PHYSICS / SCIENCE:
   - For all formulas, equations, units, and mathematical symbols: ALWAYS use LaTeX math delimiters \\( ... \\).
   - Display equations (standalone on their own line): \\[ ... \\].
   - Examples: \\( v = u + at \\), \\( F = G\\frac{{m_1 m_2}}{{r^2}} \\), \\( E = mc^2 \\), \\( \\vec{{F}} = m\\vec{{a}} \\), \\( \\lambda = \\frac{{h}}{{p}} \\), \\( 10\\ \\Omega \\), \\( 5\\ \\mu\\text{{F}} \\), \\( 9.8\\ \\text{{m/s}}^2 \\).
   - Greek letters: \\( \\alpha, \\beta, \\gamma, \\theta, \\lambda, \\mu, \\pi, \\rho, \\sigma, \\tau, \\phi, \\omega, \\Delta, \\Omega \\).

2. CHEMISTRY:
   - For all chemical formulas, equations, reactions, states, and equilibrium: ALWAYS use LaTeX \\ce{{...}} inside math delimiters \\( \\ce{{...}} \\).
   - Examples: \\( \\ce{{2H2 + O2 -> 2H2O}} \\), \\( \\ce{{CaCO3 ->[\\Delta] CaO + CO2 ^}} \\), \\( \\ce{{Fe^{{2+}} + 2e- -> Fe}} \\), \\( \\ce{{H2SO4}} \\), \\( \\ce{{CH3COOH <=> CH3COO- + H+}} \\).
   - States of matter: \\( \\ce{{NaCl (aq)}} \\), \\( \\ce{{H2O (l)}} \\), \\( \\ce{{CO2 (g)}} \\), \\( \\ce{{AgCl (s) v}} \\).

3. MATHEMATICS / STATISTICS:
   - ALWAYS use LaTeX delimiters for ALL mathematical content — expressions, fractions, integrals, summations, limits, matrices, sets, and inequalities.
   - Inline math: \\( ... \\). Display/block math: \\[ ... \\].
   - Fractions: \\( \\frac{{a}}{{b}} \\), \\( \\frac{{d}}{{dx}}(x^n) = nx^{{n-1}} \\).
   - Integrals: \\( \\int_{{a}}^{{b}} f(x)\\,dx \\), \\( \\int x^2\\,dx = \\frac{{x^3}}{{3}} + C \\).
   - Summation: \\( \\sum_{{i=1}}^{{n}} i = \\frac{{n(n+1)}}{{2}} \\).
   - Limits: \\( \\lim_{{x \\to 0}} \\frac{{\\sin x}}{{x}} = 1 \\).
   - Matrices: \\( \\begin{{pmatrix}} a & b \\\\ c & d \\end{{pmatrix}} \\).
   - Sets / Probability: \\( P(A \\cup B) \\), \\( \\binom{{n}}{{r}} \\), \\( \\bar{{x}} \\), \\( \\sigma^2 \\).
   - Geometry / Trigonometry: \\( \\sin\\theta \\), \\( \\cos^2\\theta + \\sin^2\\theta = 1 \\), \\( \\triangle ABC \\), \\( \\angle BAC = 60^\\circ \\).
   - NEVER write math as plain text (e.g. write \\( x^2 + 3x - 4 = 0 \\) NOT 'x^2 + 3x - 4 = 0').

4. ACCOUNTANCY / COMMERCE / ECONOMICS / BUSINESS STUDIES:
   - For Journal Entries, Ledger Accounts, Balance Sheets, Trial Balances, Cash Flow Statements, Trading & Profit & Loss Accounts, Ratio Analysis tables: ALWAYS use a properly formatted markdown table OR a plaintext ASCII ledger table inside a code block.
   - Journal Entry example (use markdown table):
     | Date | Particulars | L.F. | Dr (₹) | Cr (₹) |
     |------|-------------|------|--------|--------|
     | ...  | ...         | ...  | ...    | ...    |
   - Ledger Account / T-account example (use code block):
     ```
     Dr.          [Account Name] Account          Cr.
     Date | Particulars | J.F. | Amount  |  Date | Particulars | J.F. | Amount
     -----|-------------|------|---------|-------|-------------|------|-------
     ```
   - Balance Sheet / Financial Statement example (use markdown table with clear Liabilities | Amount | Assets | Amount columns).
   - Monetary amounts: ALWAYS use ₹ symbol (e.g. ₹ 50,000 or Rs. 50,000). NEVER use $ for Indian Rupees.
   - Ratios and formulae: write in plain text within the question (e.g. "Current Ratio = Current Assets / Current Liabilities") OR use LaTeX \\( ... \\) for formula display.

5. PROGRAMMING (Java, Python, C++, C, SQL, HTML, etc.):
   - For ALL multi-line code snippets, programs, and code tracing problems: ALWAYS enclose in markdown code blocks with the exact language identifier:
     ```java
     public class Solution {{
         public static void main(String[] args) {{
             System.out.println("Result: " + (5 + 3));
         }}
     }}
     ```
   - Preserve clean, accurate 4-space indentation and line breaks.
   - For inline keywords, variable names, class names, and method signatures: ALWAYS use single backticks: `public static void main`, `int x = 0;`, `Character.isLetter()`, `ArrayList<String>`, `i++`, `print()`, `SELECT * FROM`.
   - NEVER use LaTeX math mode \\( ... \\) for programming code or variable names with underscores!
   - When asked for programming questions (e.g. remarks or long answers): provide realistic, syntactically correct code snippets with proper logic and clear specifications.

CRITICAL JSON RULES:
{lang_rule_paper}
- Return ONLY a valid JSON object with the format:
{{
  "sections": [
    {{
      "section_title": "Title of the section",
      "instruction": "Section instruction",
      "marks_per_question": {1 if not is_graded else 1},
      "questions": [
        {{
          "question": "Question statement",
          "options": ["(a) ...", "(b) ...", "(c) ...", "(d) ..."],
          "sub_questions": [],
          "needs_image": false,
          "image_keywords": []
        }}
      ]
    }}
  ]
}}
- For MCQ and True/False, provide choices in "options". For other question types, "options" must be empty [].
- For Case Study, provide the context passage in "question" and sub-questions in "sub_questions".
- Return ONLY valid raw JSON without markdown fences."""

            timeout_val = 120
            import time as time_mod
            for retry in range(4):
                res_dict, prov, err = generate_ai_content(batch_prompt, timeout=timeout_val)
                if res_dict:
                    try:
                        raw_text = res_dict['candidates'][0]['content']['parts'][0]['text'].strip()
                        if raw_text.startswith('```'):
                            first_nl = raw_text.find('\n')
                            if first_nl != -1: raw_text = raw_text[first_nl + 1:].strip()
                            else: raw_text = raw_text.lstrip('`').strip()
                            if raw_text.startswith('json'): raw_text = raw_text[4:].strip()
                        if raw_text.endswith('```'):
                            raw_text = raw_text[:-3].strip()
                        parsed = safe_json_loads(raw_text)
                        if parsed and isinstance(parsed, dict) and parsed.get('sections'):
                            secs = parsed.get('sections', [])
                            if any(s.get('questions') for s in secs):
                                return secs
                        elif parsed and isinstance(parsed, list):
                            if any(s.get('questions') for s in parsed):
                                return parsed
                    except Exception as e:
                        app.logger.warning(f"Batch {b_idx} JSON parse error (attempt {retry+1}): {e}")
                
                # Backoff delay if failed or rate-limited
                wait_sec = 3 * (retry + 1)
                app.logger.info(f"Batch {b_idx} attempt {retry+1} failed ({err or 'empty'}). Retrying in {wait_sec}s...")
                time_mod.sleep(wait_sec)
            
            app.logger.error(f"Batch {b_idx} completely failed after 4 retries.")
            return []

        # Execute all batches with safe concurrency (max 2 workers to avoid HTTP 429 quota spikes)
        combined_sections = []
        if len(batches) == 1:
            combined_sections = generate_single_batch(batches[0], 0)
        else:
            import concurrent.futures, time as time_mod
            with concurrent.futures.ThreadPoolExecutor(max_workers=min(len(batches), 2)) as executor:
                future_to_idx = {}
                for i, b in enumerate(batches):
                    future_to_idx[executor.submit(generate_single_batch, b, i)] = i
                    time_mod.sleep(0.5) # Gentle stagger
                results_by_idx = {}
                for future in concurrent.futures.as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        results_by_idx[idx] = future.result()
                    except Exception as exc:
                        app.logger.error(f"Batch {idx} generated an exception: {exc}")
                        results_by_idx[idx] = []
                
                # Combine in original sequence
                for i in range(len(batches)):
                    combined_sections.extend(results_by_idx.get(i, []))

        # Merge consecutive sections of the same title (e.g. MCQ Part 1 & MCQ Part 2)
        merged_sections = []
        for sec in combined_sections:
            clean_title = re.sub(r'\s*\(Part\s*\d+\)', '', sec.get('section_title', ''), flags=re.IGNORECASE).strip()
            # Strip any LLM-prepended section prefix like "Section A: ", "Section 1 - "
            clean_title = re.sub(r'^Section\s+[A-Z0-9]+[\s:\-—]+', '', clean_title, flags=re.IGNORECASE).strip()
            sec['section_title'] = clean_title
            
            norm_key = re.sub(r'[^a-zA-Z0-9]', '', clean_title).lower()
            if merged_sections:
                prev_norm_key = re.sub(r'[^a-zA-Z0-9]', '', merged_sections[-1].get('section_title', '')).lower()
                if norm_key and norm_key == prev_norm_key:
                    merged_sections[-1]['questions'].extend(sec.get('questions', []))
                    continue

            merged_sections.append(sec)

        # Re-assign standard Section A, Section B, Section C... labels and sequential question numbers
        q_counter = 1
        for s_idx, sec in enumerate(merged_sections):
            sec['section_label'] = f"Section {chr(ord('A') + s_idx)}"
            for q in sec.get('questions', []):
                q['number'] = q_counter
                q_counter += 1

        paper_data = {'sections': merged_sections}
        if not paper_data.get('sections') or not any(s.get('questions') for s in paper_data['sections']):
            return jsonify({'status': 'error', 'message': 'Failed to generate question paper. Please check AI service or try again.'}), 500


        # Clean any raw metadata fields or trailing JSON strings from question content
        if isinstance(paper_data, dict) and 'sections' in paper_data:
            for sec in paper_data.get('sections', []):
                for q in sec.get('questions', []):
                    if isinstance(q.get('question'), str):
                        q['question'] = re.sub(r'\{"?meta(data)?"?:?.*\}', '', q['question'], flags=re.IGNORECASE).strip()
                    if isinstance(q.get('sub_questions'), list):
                        cleaned_sq = []
                        for sq in q['sub_questions']:
                            if isinstance(sq, str):
                                cleaned_sq.append(re.sub(r'\{"?meta(data)?"?:?.*\}', '', sq, flags=re.IGNORECASE).strip())
                            elif isinstance(sq, dict) and 'question' in sq and isinstance(sq['question'], str):
                                sq['question'] = re.sub(r'\{"?meta(data)?"?:?.*\}', '', sq['question'], flags=re.IGNORECASE).strip()
                                cleaned_sq.append(sq)
                            else:
                                cleaned_sq.append(sq)
                        q['sub_questions'] = cleaned_sq

        # Build logo base64 for the preview
        logo_path   = get_setting('logo_path', '')
        logo_base64 = ''
        if logo_path:
            abs_path = os.path.join(app.static_folder, logo_path)
            if os.path.exists(abs_path):
                with open(abs_path, 'rb') as f:
                    lb  = f.read()
                    ext = os.path.splitext(abs_path)[1].lower()
                    mt  = 'image/png' if ext == '.png' else 'image/jpeg'
                    logo_base64 = f"data:{mt};base64,{base64.b64encode(lb).decode()}"

        # Automatic Repository Save into paper/Class_X/Section_Y/
        teacher_name = 'Admin'
        if session.get('teacher_id'):
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT name FROM teachers WHERE id=?", (session['teacher_id'],))
            t_row = c.fetchone()
            if t_row: teacher_name = t_row['name']
            conn.close()
            
        section_val = request.form.get('section', 'A').strip() or 'A'
        abs_path, doc_filename, rel_link, *rest = save_paper_to_repository(
            paper_data, class_, section_val, subject, exam_type, teacher_name,
            meta={
                'school_name': school_nm,
                'school_address': school_addr,
                'academic_session': academic_sess,
                'duration': duration,
                'max_marks': max_marks,
                'topic': topics,
                'full_syllabus': full_syllabus,
                'logo_path': logo_path,
                'paper_type': paper_type,
                'is_graded': is_graded
            },
            paper_type=paper_type
        )
        docx_rel_link = rest[0] if rest and rest[0] else rel_link

        # Record into test_generation_history table
        teacher_id_val = str(session.get('teacher_id')) if session.get('teacher_id') else 'admin'
        try:
            conn = get_db()
            c = conn.cursor()
            clean_pt = paper_type.replace(' ', '_')
            c.execute("""INSERT INTO test_generation_history
                         (teacher_id, class, section, subject, chapter, test_no, output_mode,
                          total_questions, mcq_count, assertion_count, very_short_count,
                          short_count, long_count, case_study_count, remark, paper_type)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                      (teacher_id_val, class_, section_val, subject, topics or 'All Topics',
                       f"{clean_pt}_{exam_type}", 'docx', total_q, mcq_count, ar_count,
                       vs_count, sh_count, lg_count, cs_count, remarks, paper_type))
            conn.commit()
            conn.close()
        except Exception as ex_db:
            app.logger.warning(f"Could not record paper into test_generation_history: {ex_db}")

        return jsonify({
            'status': 'success',
            'paper':  paper_data,
            'doc_link': f"/api/shared_papers/download/{rel_link}",
            'docx_link': f"/api/shared_papers/download/{docx_rel_link}",
            'meta': {
                'paper_type':       paper_type,
                'is_graded':        is_graded,
                'school_name':      school_nm,
                'school_address':   school_addr,
                'academic_session': academic_sess,
                'exam_type':        exam_type,
                'class':            class_,
                'subject':          subject,
                'topic':            topics,
                'full_syllabus':    full_syllabus,
                'duration':         duration,
                'max_marks':        max_marks,
                'logo_base64':      logo_base64,
                'logo_path':        logo_path,
                'date':             datetime.datetime.now().strftime('%d / %m / %Y'),
            }
        })


    except urllib.error.HTTPError as e:
        return jsonify({'status': 'error', 'message': f'Gemini API: {e.read().decode()}'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/shared_papers', methods=['GET'])
def list_shared_papers():
    """Browse shared paper/Class_<N>/Section_<X>/ directory repository."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    
    cls_param = request.args.get('class', '').strip()
    sec_param = request.args.get('section', '').strip()
    
    paper_base = os.path.join(app.root_path, 'paper')
    if not os.path.exists(paper_base):
        return jsonify({'status': 'success', 'files': []})
        
    papers_list = []
    for root, dirs, files in os.walk(paper_base):
        for f in files:
            if f.endswith('.doc') or f.endswith('.docx'):
                abs_path = os.path.join(root, f)
                rel_path = os.path.relpath(abs_path, paper_base).replace('\\', '/')
                path_parts = rel_path.split('/')
                
                # Check path filters
                file_cls = path_parts[0].replace('Class_', '') if len(path_parts) > 1 else ''
                file_sec = path_parts[1].replace('Section_', '') if len(path_parts) > 2 else ''
                
                if cls_param and file_cls.lower() != cls_param.lower():
                    continue
                if sec_param and file_sec.lower() != sec_param.lower():
                    continue
                    
                stat = os.stat(abs_path)
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                size_kb = round(stat.st_size / 1024, 1)
                
                papers_list.append({
                    'filename': f,
                    'class': file_cls,
                    'section': file_sec,
                    'rel_path': rel_path,
                    'size_kb': size_kb,
                    'modified_at': mtime,
                    'download_url': f"/api/shared_papers/download/{rel_path}"
                })
                
    papers_list.sort(key=lambda x: x['modified_at'], reverse=True)
    return jsonify({'status': 'success', 'files': papers_list})


@app.route('/api/shared_papers/download/<path:rel_path>', methods=['GET'])
def download_shared_paper(rel_path):
    """Download .doc paper file from repository."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    
    if rel_path.startswith('paper/') or rel_path.startswith('paper\\'):
        rel_path = rel_path[6:]

    paper_base = os.path.join(app.root_path, 'paper')
    abs_file_path = os.path.abspath(os.path.join(paper_base, rel_path))
    
    # Security check: prevent path traversal out of paper directory
    if not abs_file_path.startswith(os.path.abspath(paper_base)):
        return jsonify({'status': 'error', 'message': 'Invalid file path'}), 400
        
    if not os.path.exists(abs_file_path):
        return jsonify({'status': 'error', 'message': 'File not found'}), 404
        
    filename = os.path.basename(abs_file_path)
    mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' if filename.endswith('.docx') else 'application/msword'
    return send_file(abs_file_path, as_attachment=True, download_name=filename, mimetype=mimetype)


@app.route('/api/shared_papers/delete', methods=['POST'])
def delete_shared_paper():
    """API endpoint to delete a shared paper file from the repository."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
    
    data = request.get_json() if request.is_json else request.form
    rel_path = (data.get('rel_path') or '').strip()
    if not rel_path:
        return jsonify({'status': 'error', 'message': 'Missing file path'}), 400

    if rel_path.startswith('paper/') or rel_path.startswith('paper\\'):
        rel_path = rel_path[6:]

    paper_base = os.path.join(app.root_path, 'paper')
    abs_file_path = os.path.abspath(os.path.join(paper_base, rel_path))
    
    # Path traversal protection
    if not abs_file_path.startswith(os.path.abspath(paper_base)):
        return jsonify({'status': 'error', 'message': 'Invalid file path'}), 400

    if not os.path.exists(abs_file_path):
        return jsonify({'status': 'error', 'message': 'File not found'}), 404

    try:
        os.remove(abs_file_path)
        log_audit_event('user', session.get('username', 'user'), 'DELETE_SHARED_PAPER', 'paper', rel_path, request.remote_addr)
        return jsonify({'status': 'success', 'message': 'Paper file deleted successfully.'})
    except Exception as e:
        app.logger.error(f"Error deleting shared paper file: {e}")
        return jsonify({'status': 'error', 'message': f'Failed to delete file: {e}'}), 500


# ─── CURRICULUM API (CLASSES 1 TO 12 — ALL SUBJECTS & TOPICS) ─────────────
@app.route('/api/curriculum/boards', methods=['GET'])
def api_curriculum_boards():
    """Return available education boards (CBSE, ICSE, State Board)."""
    return jsonify({
        'status': 'success',
        'boards': curriculum_data.get_available_boards()
    })


@app.route('/api/curriculum/classes', methods=['GET'])
def api_curriculum_classes():
    """Return list of supported classes 1 to 12."""
    board = request.args.get('board', 'CBSE')
    return jsonify({
        'status': 'success',
        'classes': curriculum_data.get_classes_for_board(board)
    })


@app.route('/api/curriculum/subjects', methods=['GET'])
def api_curriculum_subjects():
    """Return list of subjects for a given class and board."""
    cls = request.args.get('class', '10')
    board = request.args.get('board', 'CBSE')
    subjects = curriculum_data.get_subjects_for_class(cls, board)
    return jsonify({
        'status': 'success',
        'class': cls,
        'board': board,
        'subjects': subjects
    })


@app.route('/api/curriculum/topics', methods=['GET'])
def api_curriculum_topics():
    """Return list of chapters and topic breakdown for a given class, subject, and board."""
    cls = request.args.get('class', '10')
    subj = request.args.get('subject', 'Science')
    board = request.args.get('board', 'CBSE')
    chapters = curriculum_data.get_chapters_and_topics(cls, subj, board)
    return jsonify({
        'status': 'success',
        'class': cls,
        'subject': subj,
        'board': board,
        'chapters': chapters
    })


@app.route('/api/curriculum/search', methods=['GET'])
def api_curriculum_search():
    """Search topics or chapters across all classes and subjects."""
    q = request.args.get('q', '').strip()
    results = curriculum_data.search_curriculum(q)
    return jsonify({
        'status': 'success',
        'query': q,
        'count': len(results),
        'results': results
    })


@app.route('/admin/check_class_teacher', methods=['GET'])
@admin_required
def admin_check_class_teacher():
    """Check if Class Teacher is already assigned for Class + Section."""
    cls = request.args.get('class', '').strip()
    sec = request.args.get('section', '').strip()
    t_id = request.args.get('teacher_id', type=int)
    
    if not cls:
        return jsonify({'has_ct': False})
        
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT ta.id, t.name FROM teacher_assignments ta
                 JOIN teachers t ON ta.teacher_id = t.id
                 WHERE ta.class=? AND (ta.section=? OR ta.section='') AND ta.is_class_teacher=1
                 AND (? IS NULL OR ta.teacher_id != ?)""",
              (cls, sec, t_id, t_id))
    ct = c.fetchone()
    conn.close()
    
    if ct:
        return jsonify({'has_ct': True, 'ct_name': ct['name']})
    return jsonify({'has_ct': False})


@app.route('/api/teacher/mcq_history', methods=['GET'])
def get_teacher_mcq_history():
    """Get MCQ test generation history for teachers, including pre-existing test papers."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401
        
    teacher_id = session.get('teacher_id')
    conn = get_db()
    c = conn.cursor()
    
    # 1. Fetch from mcq_test_history
    if teacher_id:
        c.execute("""SELECT * FROM mcq_test_history WHERE teacher_id=? ORDER BY id DESC""", (teacher_id,))
    else:
        c.execute("""SELECT * FROM mcq_test_history ORDER BY id DESC""")
    rows = [dict(r) for r in c.fetchall()]
    
    # 2. Also fetch pre-existing tests from test_papers table
    if teacher_id:
        c.execute("""SELECT tp.id, tp.uploaded_by as teacher_id, COALESCE(t.name, 'Faculty') as teacher_name, 
                            tp.class, tp.section, tp.subject, tp.test_no, tp.question_count, tp.created_at
                     FROM test_papers tp
                     LEFT JOIN teachers t ON tp.uploaded_by = CAST(t.id AS TEXT)
                     WHERE tp.uploaded_by = ? OR tp.uploader_type='admin'
                     ORDER BY tp.created_at DESC""", (str(teacher_id),))
    else:
        c.execute("""SELECT tp.id, tp.uploaded_by as teacher_id, COALESCE(t.name, 'Admin') as teacher_name, 
                            tp.class, tp.section, tp.subject, tp.test_no, tp.question_count, tp.created_at
                     FROM test_papers tp
                     LEFT JOIN teachers t ON tp.uploaded_by = CAST(t.id AS TEXT)
                     ORDER BY tp.created_at DESC""")
    tp_rows = [dict(r) for r in c.fetchall()]
    
    # Combine & deduplicate by (class, section, subject, test_no)
    seen_keys = set()
    combined = []
    for r in rows:
        key = (str(r.get('class')), str(r.get('section')), str(r.get('subject')).lower(), str(r.get('test_no')))
        seen_keys.add(key)
        combined.append(r)
        
    for r in tp_rows:
        key = (str(r.get('class')), str(r.get('section')), str(r.get('subject')).lower(), str(r.get('test_no')))
        if key not in seen_keys:
            seen_keys.add(key)
            combined.append(r)
            
    conn.close()
    return jsonify({'status': 'success', 'history': combined})


@app.route('/api/csv_template')
def download_csv_template():
    """FEAT-001: Downloadable CSV template for question paper upload."""
    import io
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Questions', 'option a', 'option b', 'option c', 'option d',
                     'correct option', 'question_type', 'marks', 'negative_mark'])
    writer.writerow([
        'What is the speed of light in vacuum?',
        '3 x 10^8 m/s', '3 x 10^6 m/s', '3 x 10^10 m/s', '3 x 10^4 m/s',
        'option a', 'MCQ', '1', '0.33'
    ])
    writer.writerow([
        'Assertion: Water boils at 100°C. Reason: Boiling point depends on atmospheric pressure.',
        'Both A and R are true and R is the correct explanation of A',
        'Both A and R are true but R is NOT the correct explanation of A',
        'A is true but R is false',
        'A is false but R is true',
        'option_a', 'Assertion-Reason', '1', '0.33'
    ])
    writer.writerow([
        'Explain the process of photosynthesis.', '', '', '', '',
        'N/A', 'Short', '3', '0'
    ])
    writer.writerow([
        'Define Newton\'s first law of motion and give two examples.', '', '', '', '',
        'N/A', 'Very Short', '2', '0'
    ])
    writer.writerow([
        'Discuss the structure and functions of DNA.', '', '', '', '',
        'N/A', 'Long', '5', '0'
    ])
    output.seek(0)
    from flask import Response
    return Response(
        output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': 'attachment; filename=question_upload_template.csv'}
    )

@app.route('/announcement/<int:bid>')
def announcement_permalink(bid):
    """FEAT-005: Permanent link to a single announcement."""
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT b.*, COALESCE(t.name,'Admin') as poster_name
                 FROM bulletins b LEFT JOIN teachers t ON b.posted_by=t.id AND b.poster_type='teacher'
                 WHERE b.id=? AND b.is_active=1""", (bid,))
    bulletin = c.fetchone()
    conn.close()
    if not bulletin:
        return "Announcement not found or has been removed.", 404
    bulletin = dict(bulletin)
    school_name = get_setting('school_name', 'RRB Group of Schools')
    logo_path   = get_setting('logo_path', '')
    return render_template('announcement_permalink.html',
                           bulletin=bulletin, school_name=school_name, logo_path=logo_path)

@app.route('/api/shared_tests')
def api_shared_tests():
    """FEAT-008: List all shared/available tests for a class+subject."""
    class_   = request.args.get('class','').strip()
    subject  = request.args.get('subject','').strip()
    if not class_ or not subject:
        return jsonify({'tests': []})
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT tp.id, tp.filename, tp.test_no, tp.section, tp.question_count,
                        tp.is_active, tp.created_at, tp.uploader_type,
                        COALESCE(t.name,'Admin') as creator_name
                 FROM test_papers tp
                 LEFT JOIN teachers t ON tp.uploaded_by=t.id AND tp.uploader_type='teacher'
                 WHERE tp.class=? AND tp.subject=?
                 ORDER BY tp.created_at DESC""", (class_, subject))
    tests = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'tests': tests})

@app.route('/api/assign_existing_test', methods=['POST'])
def assign_existing_test():
    """FEAT-008: Link an existing test to a class/section without duplicating questions."""
    if not session.get('teacher_logged_in') and not session.get('admin_logged_in'):
        return jsonify({'status':'error'}), 401
    paper_id  = request.form.get('paper_id', type=int)
    section   = request.form.get('section','').strip()
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM test_papers WHERE id=?", (paper_id,))
    paper = c.fetchone()
    if not paper:
        conn.close()
        return jsonify({'status':'error','message':'Test not found'}), 404
    # Update section assignment (or leave as shared)
    c.execute("UPDATE test_papers SET section=?, is_active=1 WHERE id=?", (section, paper_id))
    conn.commit()
    conn.close()
    return jsonify({'status':'success','message':f'Test "{paper["filename"]}" assigned to Section {section or "All"}'})


@app.route('/api/token', methods=['POST'])
@limiter.limit("20 per minute")
def api_issue_token():
    """
    Issue JWT access token for Admin, Teacher, or Student API authentication.
    Accepts JSON or Form Data payload with:
      - role: 'admin', 'teacher', or 'student'
      - username / mobile / student_id / roll_no
      - password (for admin & teacher)
      - name, class, section (for student)
    """
    data = request.get_json(silent=True) or request.form or {}
    role = (data.get('role') or 'admin').strip().lower()
    
    if role == 'admin':
        username = (data.get('username') or data.get('identity') or 'admin').strip()
        password = data.get('password') or ''
        
        admin_pass = os.environ.get('ADMIN_PASSWORD', 'admin123').strip()
        if password == admin_pass:
            token, exp_ts = generate_jwt_token(identity=username, role='admin', name='System Admin')
            return jsonify({
                'status': 'success',
                'access_token': token,
                'token_type': 'Bearer',
                'expires_in': JWT_EXPIRATION_HOURS * 3600,
                'user': {'id': username, 'name': 'System Admin', 'role': 'admin'}
            })
        else:
            return jsonify({'status': 'error', 'message': 'Invalid admin credentials'}), 401
            
    elif role == 'teacher':
        mobile = (data.get('mobile') or data.get('username') or data.get('identity') or '').strip()
        password = data.get('password') or ''
        
        if not mobile or not password:
            return jsonify({'status': 'error', 'message': 'Mobile and password required'}), 400
            
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM teachers WHERE mobile=? AND status='active'", (mobile,))
        teacher = c.fetchone()
        conn.close()
        
        if teacher and verify_password(password, teacher['password']):
            token, exp_ts = generate_jwt_token(identity=teacher['id'], role='teacher', name=teacher['name'])
            return jsonify({
                'status': 'success',
                'access_token': token,
                'token_type': 'Bearer',
                'expires_in': JWT_EXPIRATION_HOURS * 3600,
                'user': {'id': teacher['id'], 'name': teacher['name'], 'role': 'teacher', 'mobile': teacher['mobile']}
            })
        else:
            return jsonify({'status': 'error', 'message': 'Invalid teacher mobile or password'}), 401
            
    elif role == 'student':
        student_id = (data.get('student_id') or data.get('roll_no') or data.get('username') or data.get('identity') or '').strip()
        name = (data.get('name') or 'Student').strip()
        class_ = (data.get('class') or '').strip()
        section = (data.get('section') or '').strip()
        
        if not student_id:
            return jsonify({'status': 'error', 'message': 'Student ID / Roll No required'}), 400
            
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
        student = c.fetchone()
        conn.close()
        
        student_name = student['name'] if student else name
        token, exp_ts = generate_jwt_token(identity=student_id, role='student', name=student_name)
        return jsonify({
            'status': 'success',
            'access_token': token,
            'token_type': 'Bearer',
            'expires_in': JWT_EXPIRATION_HOURS * 3600,
            'user': {
                'id': student_id,
                'name': student_name,
                'role': 'student',
                'class': student['class'] if student else class_,
                'section': student['section'] if student else section
            }
        })
    else:
        return jsonify({'status': 'error', 'message': f'Invalid role: {role}. Must be admin, teacher, or student.'}), 400


@app.route('/api/user/me', methods=['GET'])
@jwt_required()
def api_get_current_user():
    """Returns details of the currently authenticated JWT user context."""
    return jsonify({
        'status': 'success',
        'user': g.jwt_user,
        'payload': g.jwt_payload
    })


# ── NCERT EXEMPLAR BANK API ENDPOINTS ─────────────────────────────────────────
@app.route('/api/exemplar/chapters', methods=['GET'])
def api_exemplar_chapters():
    """List available NCERT Exemplar chapters in the local bank."""
    subject = request.args.get('subject')
    class_num = request.args.get('class')
    chapters = exemplar_extractor.get_exemplar_chapters(subject=subject, class_num=class_num)
    return jsonify({'status': 'success', 'chapters': chapters, 'total': len(chapters)})


@app.route('/api/exemplar/upload', methods=['POST'])
def api_exemplar_upload():
    """Upload and process an NCERT Exemplar chapter PDF file into the tagged local bank."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401

    pdf_file = request.files.get('pdf_file')
    if not pdf_file or not pdf_file.filename:
        return jsonify({'status': 'error', 'message': 'Please select a PDF file'}), 400

    if not pdf_file.filename.lower().endswith('.pdf'):
        return jsonify({'status': 'error', 'message': 'Only PDF files are supported'}), 400

    subject = request.form.get('subject', 'General').strip()
    class_num = request.form.get('class', '10').strip()
    chapter_name = request.form.get('chapter', '').strip() or pdf_file.filename.rsplit('.', 1)[0]

    upload_dir = os.path.join(app.root_path, 'uploads', 'exemplar_pdfs')
    os.makedirs(upload_dir, exist_ok=True)
    temp_pdf_path = os.path.join(upload_dir, f"exemplar_{class_num}_{subject}_{chapter_name}.pdf")
    pdf_file.save(temp_pdf_path)

    try:
        chapter_entry = exemplar_extractor.process_exemplar_pdf(
            pdf_path=temp_pdf_path,
            subject=subject,
            class_num=class_num,
            chapter_name=chapter_name
        )
        return jsonify({
            'status': 'success',
            'message': f"Extracted {chapter_entry.get('total_questions', 0)} questions and {chapter_entry.get('total_images', 0)} images.",
            'data': chapter_entry
        })
    except Exception as e:
        app.logger.error(f"Exemplar extraction error: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/exemplar/generate_paper', methods=['POST'])
def api_exemplar_generate_paper():
    """Generate a test paper directly from the local NCERT Exemplar bank."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401

    data = request.get_json(silent=True) or request.form.to_dict()
    subject = data.get('subject', '').strip()
    class_num = data.get('class', '').strip()
    chapter_name = data.get('chapter', '').strip()
    mcq_count = int(data.get('mcq_count', 10) or 10)
    short_count = int(data.get('short_count', 5) or 5)
    long_count = int(data.get('long_count', 3) or 3)

    paper_data = exemplar_extractor.generate_paper_from_exemplar(
        subject=subject,
        class_num=class_num,
        chapter_name=chapter_name if chapter_name else None,
        mcq_count=mcq_count,
        short_count=short_count,
        long_count=long_count
    )

    if not paper_data.get('sections') or not any(s.get('questions') for s in paper_data['sections']):
        return jsonify({
            'status': 'error',
            'message': f"No questions found in Exemplar Bank for {subject} Class {class_num}."
        }), 404

    return jsonify({
        'status': 'success',
        'provider': 'NCERT Exemplar Bank',
        'data': paper_data,
        'paper_data': paper_data
    })


# ---------------------------------------------------------------------------
# /theme.css — Dynamic CSS custom properties from settings table
# ---------------------------------------------------------------------------
@app.route('/theme.css')
def serve_theme_css():
    """Serves dynamic CSS with :root variables from the settings table."""
    cfg = get_site_config()
    primary   = cfg['theme_primary']
    secondary = cfg['theme_secondary']
    accent    = cfg['theme_accent']
    student   = cfg['theme_student']
    font_web  = cfg['theme_font_web']

    def _darken(hex_col, factor=0.75):
        hex_col = hex_col.lstrip('#')
        try:
            r, g, b = int(hex_col[0:2], 16), int(hex_col[2:4], 16), int(hex_col[4:6], 16)
            return '#{:02x}{:02x}{:02x}'.format(int(r*factor), int(g*factor), int(b*factor))
        except Exception:
            return '#0d1457'

    css_font = f"'{font_web}', sans-serif" if font_web in _GOOGLE_FONTS_MAP else f'{font_web}, sans-serif'
    css = f""":root {{
  --primary:        {primary};
  --primary-light:  {secondary};
  --primary-lighter:{secondary};
  --primary-dark:   {_darken(primary)};
  --accent:         {accent};
  --green-dk:       {student};
  --green:          {student};
  --theme-font:     {css_font};
}}
body {{ font-family: {css_font}; }}
"""
    from flask import Response
    return Response(css, mimetype='text/css',
                    headers={'Cache-Control': 'no-store'})


# ---------------------------------------------------------------------------
# /admin/save_theme — Save white-label theme, font, and developer credit settings
# ---------------------------------------------------------------------------
@app.route('/admin/save_theme', methods=['POST'])
def admin_save_theme():
    """Save white-label theme, font, credit, and active language settings."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized. Please log in as Admin.'}), 403
    data = request.get_json(silent=True) or request.form
    allowed_keys = [
        'theme_primary', 'theme_secondary', 'theme_accent', 'theme_student',
        'theme_icon_color', 'theme_btn_bg', 'theme_btn_text',
        'theme_heading_color', 'theme_body_text', 'theme_muted_text',
        'theme_page_bg', 'theme_card_bg', 'theme_sidebar_bg',
        'theme_link_color', 'theme_hover_color', 'theme_border_color',
        'theme_success_color', 'theme_warning_color', 'theme_danger_color',
        'theme_font_web', 'theme_font_export', 'developer_credit', 'active_language'
    ]
    for key in allowed_keys:
        if key in data:
            val = str(data[key]).strip()
            if val or key in ('developer_credit',):
                set_setting(key, val)
    return jsonify({'status': 'success', 'message': 'Theme settings saved successfully.'})


@app.route('/api/set_language', methods=['POST'])
def api_set_language():
    """Sets student language in session for instant UI switching."""
    data = request.get_json(silent=True) or request.form or {}
    lang = (data.get('lang_code') or 'en').strip().lower()[:10]
    session['student_lang'] = lang
    return jsonify({'status': 'success', 'active_language': lang})


# ---------------------------------------------------------------------------
# /admin/translate — Trigger AI translation for a language (background thread)
# ---------------------------------------------------------------------------
import threading, hashlib, time as _time

def _run_translation(lang_code):
    """Background worker: translates all UI strings via Gemini and caches them."""
    try:
        from ui_strings import UI_STRINGS
    except Exception as e:
        app.logger.error(f"Failed to import UI_STRINGS: {e}")
        try:
            with get_db() as _conn:
                _conn.execute(
                    "UPDATE translation_meta SET status='error', last_updated=CURRENT_TIMESTAMP WHERE lang_code=?",
                    (lang_code,)
                )
        except Exception:
            pass
        return

    gemini_key = _get_gemini_key()
    if not gemini_key:
        app.logger.error("Gemini API key not configured for translation")
        try:
            with get_db() as _conn:
                _conn.execute(
                    "UPDATE translation_meta SET status='error', last_updated=CURRENT_TIMESTAMP WHERE lang_code=?",
                    (lang_code,)
                )
        except Exception:
            pass
        return

    keys   = list(UI_STRINGS.keys())
    total  = len(keys)
    source_hash = hashlib.sha256(
        '|'.join(f'{k}={v}' for k, v in UI_STRINGS.items()).encode()
    ).hexdigest()

    conn = get_db()
    try:
        conn.execute(
            """INSERT OR REPLACE INTO translation_meta
               (lang_code, source_hash, status, progress, total_strings, translated_count, last_updated)
               VALUES (?, ?, 'in_progress', 0, ?, 0, CURRENT_TIMESTAMP)""",
            (lang_code, source_hash, total)
        )
        conn.commit()

        BATCH = 20
        translated_count = 0
        model_name = get_working_model(gemini_key) or 'gemini-1.5-flash-latest'
        lang_label_map = {
            'hi': 'Hindi', 'mr': 'Marathi', 'ta': 'Tamil',
            'te': 'Telugu', 'bn': 'Bengali', 'gu': 'Gujarati',
            'kn': 'Kannada', 'ml': 'Malayalam', 'pa': 'Punjabi',
        }
        lang_label = lang_label_map.get(lang_code, lang_code)

        for batch_start in range(0, total, BATCH):
            batch_keys = keys[batch_start:batch_start + BATCH]
            batch_vals = {k: UI_STRINGS[k] for k in batch_keys}

            prompt = (
                f"You are a professional educational software translator. "
                f"Translate these UI strings from English to {lang_label}. "
                f"Return ONLY a JSON object with the same keys, values translated. "
                f"Keep placeholders like {{{{name}}}} unchanged. "
                f"JSON to translate:\n{json.dumps(batch_vals, ensure_ascii=False)}"
            )
            try:
                import urllib.request, urllib.error
                req_body = json.dumps({
                    'contents': [{'parts': [{'text': prompt}]}],
                    'generationConfig': {'temperature': 0.1, 'maxOutputTokens': 2048}
                }).encode()
                url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
                       f"{model_name}:generateContent?key={gemini_key}")
                req = urllib.request.Request(url, data=req_body,
                                             headers={'Content-Type': 'application/json'})
                with urllib.request.urlopen(req, timeout=60) as resp:
                    result = json.loads(resp.read().decode())
                raw_text = result['candidates'][0]['content']['parts'][0]['text']
                # Strip markdown code fences if present
                raw_text = raw_text.strip()
                if raw_text.startswith('```'):
                    raw_text = raw_text.split('\n', 1)[1].rsplit('```', 1)[0]
                translated_batch = json.loads(raw_text)
                for k, v in translated_batch.items():
                    if k in batch_vals:
                        conn.execute(
                            "INSERT OR REPLACE INTO translation_cache (lang_code, string_key, translated) VALUES (?,?,?)",
                            (lang_code, k, v)
                        )
                translated_count += len(translated_batch)
                conn.commit()
            except Exception as e:
                app.logger.warning(f"Translation batch error ({lang_code}): {e}")

            progress = min(99, int((batch_start + BATCH) / total * 100))
            conn.execute(
                """UPDATE translation_meta SET progress=?, translated_count=?, last_updated=CURRENT_TIMESTAMP
                   WHERE lang_code=?""",
                (progress, translated_count, lang_code)
            )
            conn.commit()
            _time.sleep(1)  # Respect Gemini rate limits

        # Mark done
        conn.execute(
            """UPDATE translation_meta SET status='done', progress=100, translated_count=?,
               source_hash=?, last_updated=CURRENT_TIMESTAMP WHERE lang_code=?""",
            (translated_count, source_hash, lang_code)
        )
        conn.commit()
    except Exception as e:
        app.logger.error(f"Fatal error in _run_translation ({lang_code}): {e}")
        try:
            conn.execute(
                "UPDATE translation_meta SET status='error', last_updated=CURRENT_TIMESTAMP WHERE lang_code=?",
                (lang_code,)
            )
            conn.commit()
        except Exception:
            pass
    finally:
        try:
            conn.close()
        except Exception:
            pass


@app.route('/admin/translate', methods=['POST'])
def admin_translate():
    """Kick off background translation for a language."""
    if not session.get('admin_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403
    lang_code = request.json.get('lang_code', '').strip().lower()[:10]
    if not lang_code:
        return jsonify({'status': 'error', 'message': 'lang_code required'}), 400

    # Check if already in progress
    conn = get_db()
    row = conn.execute(
        "SELECT status FROM translation_meta WHERE lang_code=?", (lang_code,)
    ).fetchone()
    if row and row[0] == 'in_progress':
        conn.close()
        return jsonify({'status': 'already_running', 'message': 'Translation already in progress'})

    # Check if already done + hash unchanged
    try:
        from ui_strings import UI_STRINGS
        current_hash = hashlib.sha256(
            '|'.join(f'{k}={v}' for k, v in UI_STRINGS.items()).encode()
        ).hexdigest()
        if row and row[0] == 'done':
            meta = conn.execute(
                "SELECT source_hash FROM translation_meta WHERE lang_code=?", (lang_code,)
            ).fetchone()
            if meta and meta[0] == current_hash:
                conn.close()
                # Already translated, just activate
                set_setting('active_language', lang_code)
                return jsonify({'status': 'cached', 'message': 'Using existing translation'})
    except Exception:
        pass
    conn.close()

    threading.Thread(target=_run_translation, args=(lang_code,), daemon=True).start()
    return jsonify({'status': 'started', 'message': f'Translation started for {lang_code}'})


@app.route('/admin/translation_progress')
def admin_translation_progress():
    """Poll translation progress for a language."""
    if not session.get('admin_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403
    lang_code = request.args.get('lang', 'en')
    conn = get_db()
    row = conn.execute(
        "SELECT status, progress, total_strings, translated_count FROM translation_meta WHERE lang_code=?",
        (lang_code,)
    ).fetchone()
    conn.close()
    if not row:
        return jsonify({'status': 'pending', 'progress': 0, 'total': 0, 'done': 0})
    status, progress, total, done = row
    if status == 'done':
        set_setting('active_language', lang_code)
    return jsonify({'status': status, 'progress': progress, 'total': total, 'done': done})


# ---------------------------------------------------------------------------
# /admin/reset_settings — Reset all white-label settings to factory defaults
# ---------------------------------------------------------------------------
_DEFAULT_SETTINGS = {
    'school_name':       'RRB Group of Schools',
    'school_address':    '',
    'academic_session':  '',
    'logo_path':         '',
    # Brand Colors
    'theme_primary':     '#1a237e',
    'theme_secondary':   '#3949ab',
    'theme_accent':      '#10b981',
    'theme_student':     '#1b5e20',
    # Extended Colors
    'theme_icon_color':  '#4f46e5',
    'theme_btn_bg':      '#1a237e',
    'theme_btn_text':    '#ffffff',
    'theme_heading_color':   '#1e1b4b',
    'theme_body_text':       '#1e293b',
    'theme_muted_text':      '#64748b',
    'theme_page_bg':         '#f0f4f8',
    'theme_card_bg':         '#ffffff',
    'theme_sidebar_bg':      '#1a237e',
    'theme_link_color':      '#3949ab',
    'theme_hover_color':     '#4f46e5',
    'theme_border_color':    '#e2e8f0',
    'theme_success_color':   '#16a34a',
    'theme_warning_color':   '#f59e0b',
    'theme_danger_color':    '#dc2626',
    # Fonts
    'theme_font_web':    'Poppins',
    'theme_font_export': 'Arial',
    'developer_credit':  'Developed by Gaurav Shukla &amp; Team',
    'active_language':   'en',
}

_DEFAULT_THEME_SETTINGS = {
    # Brand Colors
    'theme_primary':         '#1a237e',
    'theme_secondary':       '#3949ab',
    'theme_accent':          '#10b981',
    'theme_student':         '#1b5e20',
    # Extended Colors
    'theme_icon_color':      '#4f46e5',
    'theme_btn_bg':          '#1a237e',
    'theme_btn_text':        '#ffffff',
    'theme_heading_color':   '#1e1b4b',
    'theme_body_text':       '#1e293b',
    'theme_muted_text':      '#64748b',
    'theme_page_bg':         '#f0f4f8',
    'theme_card_bg':         '#ffffff',
    'theme_sidebar_bg':      '#1a237e',
    'theme_link_color':      '#3949ab',
    'theme_hover_color':     '#4f46e5',
    'theme_border_color':    '#e2e8f0',
    'theme_success_color':   '#16a34a',
    'theme_warning_color':   '#f59e0b',
    'theme_danger_color':    '#dc2626',
    # Fonts
    'theme_font_web':        'Poppins',
    'theme_font_export':     'Arial',
    'developer_credit':      'Developed by Gaurav Shukla &amp; Team',
}

@app.route('/admin/reset_theme', methods=['POST'])
def admin_reset_theme():
    """Revert only theme colors, fonts, and developer credit to factory defaults."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized. Please log in as Admin.'}), 403
    for key, val in _DEFAULT_THEME_SETTINGS.items():
        set_setting(key, val)
    return jsonify({'status': 'success', 'message': 'Theme reset to factory defaults successfully.'})

@app.route('/admin/reset_settings', methods=['POST'])
def admin_reset_settings():
    """Revert all white-label settings to factory defaults in one click."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized. Please log in as Admin.'}), 403
    for key, val in _DEFAULT_SETTINGS.items():
        set_setting(key, val)
    return jsonify({'status': 'success', 'message': 'All settings reset to defaults.'})


# ---------------------------------------------------------------------------
# Coding Test Module Execution & Evaluation Endpoints
# ---------------------------------------------------------------------------
from coding_runner import run_code, evaluate_submission, check_compiler_availability

@app.route('/api/test_coding_runner', methods=['POST'])
def api_test_coding_runner():
    """
    Test endpoint for executing and evaluating student code snippets.
    Supports Python, Java, C, and C++.
    """
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in') and not session.get('student_id'):
        return jsonify({'status': 'error', 'message': 'Authentication required'}), 401
        
    data = request.get_json(silent=True) or request.form
    code = data.get('code', '')
    language = data.get('language', 'python')
    stdin_input = data.get('stdin_input', '')
    timeout_seconds = int(data.get('timeout_seconds', 5))
    test_cases = data.get('test_cases', [])
    
    if not code:
        return jsonify({'status': 'error', 'message': 'Code snippet is required'}), 400
        
    if test_cases and isinstance(test_cases, list):
        eval_result = evaluate_submission(code, language, test_cases, timeout_seconds=timeout_seconds)
        return jsonify(eval_result)
    else:
        run_result = run_code(code, language, stdin_input, timeout_seconds=timeout_seconds)
        return jsonify(run_result)


# ---------------------------------------------------------------------------
# Coding Test Module AI Generation & Teacher Review Endpoints
# ---------------------------------------------------------------------------

@app.route('/api/generate_coding_questions', methods=['POST'])
def api_generate_coding_questions():
    """Generate programming coding questions & test cases via Gemini AI and save as unapproved drafts."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403

    conn = None
    try:
        board = request.form.get('board', 'CBSE').strip() or 'CBSE'
        class_ = request.form.get('class', '').strip()
        section = request.form.get('section', '').strip()
        subject = request.form.get('subject', '').strip()
        topic = request.form.get('chapter', '').strip() or request.form.get('topic', '').strip() or 'Programming Fundamentals'
        test_no = request.form.get('test_no', '').strip() or f"CODING_{uuid.uuid4().hex[:6].upper()}"
        
        try:
            coding_count = int(request.form.get('coding_count', 5) or 5)
        except Exception:
            coding_count = 5

        try:
            coding_marks = float(request.form.get('coding_marks', 10) or 10)
        except Exception:
            coding_marks = 10.0

        coding_difficulty = request.form.get('coding_difficulty', 'medium').strip()
        
        allowed_langs_list = request.form.getlist('allowed_languages')
        if not allowed_langs_list:
            allowed_langs_list = ['python', 'java', 'c', 'cpp']
        allowed_langs = ','.join(allowed_langs_list)

        prompt = f"""You are an expert Computer Science teacher and coding problem author for {board} Class {class_} ({subject}).
Generate EXACTLY {coding_count} practical programming/coding problems for Topic/Syllabus: "{topic}".
Difficulty Level: {coding_difficulty.upper()}.
Target Languages: {allowed_langs}.

Return ONLY a valid JSON array of {coding_count} problem objects. DO NOT include markdown backticks around JSON if possible, or format strictly as ```json ... ```.

Each object MUST have the following structure:
[
  {{
    "title": "Short descriptive title of problem",
    "problem_statement": "Clear detailed statement of the problem, background scenario, and requirement.",
    "input_format": "Description of input format expected from stdin",
    "output_format": "Description of output format expected on stdout",
    "constraints": "Space/Time and Input bounds constraints e.g. 1 <= N <= 1000",
    "sample_input": "First sample input string",
    "sample_output": "First sample output string",
    "test_cases": [
      {{"input": "Sample input 1", "expected_output": "Expected output 1", "is_hidden": 0, "marks_weight": 1.0}},
      {{"input": "Sample input 2", "expected_output": "Expected output 2", "is_hidden": 0, "marks_weight": 1.0}},
      {{"input": "Hidden test input 1", "expected_output": "Hidden expected output 1", "is_hidden": 1, "marks_weight": 1.0}},
      {{"input": "Hidden test input 2", "expected_output": "Hidden expected output 2", "is_hidden": 1, "marks_weight": 1.0}}
    ]
  }}
]
"""
        res_dict, provider_used, err = generate_ai_content(prompt, timeout=90)
        if not res_dict:
            return jsonify({'status': 'error', 'message': f'AI Generation Error: {err or "No response from AI generator"}'}), 500

        raw_res = ""
        try:
            raw_res = res_dict['candidates'][0]['content']['parts'][0]['text'].strip()
        except Exception:
            raw_res = str(res_dict)

        parsed = safe_json_loads(raw_res)
        if isinstance(parsed, dict):
            for k in ['coding_questions', 'questions', 'problems', 'data', 'items']:
                if isinstance(parsed.get(k), list):
                    parsed = parsed[k]
                    break
            if isinstance(parsed, dict):
                parsed = [parsed]

        if not parsed or not isinstance(parsed, list):
            raw_problems = re.findall(r'\{\s*"title"\s*:\s*"(.*?)"\s*,\s*"problem_statement"\s*:\s*"(.*?)"', raw_res, re.DOTALL)
            if raw_problems:
                parsed = []
                for t_match, p_match in raw_problems:
                    parsed.append({
                        "title": t_match.strip(),
                        "problem_statement": p_match.strip(),
                        "sample_input": "",
                        "sample_output": "",
                        "test_cases": []
                    })

        if not parsed or not isinstance(parsed, list):
            return jsonify({'status': 'error', 'message': 'Failed to parse AI response into structured coding problems. Please try again.'}), 500

        uploader_type = 'admin' if session.get('admin_logged_in') else 'teacher'
        uploaded_by = 'admin' if uploader_type == 'admin' else str(session.get('teacher_id') or 'teacher')

        conn = get_db()
        c = conn.cursor()
        inserted_ids = []

        for q in parsed:
            if not isinstance(q, dict) or not q.get('problem_statement'):
                continue
            title = q.get('title', 'Coding Challenge').strip()
            statement = q.get('problem_statement', '').strip()
            inp_fmt = q.get('input_format', '').strip()
            out_fmt = q.get('output_format', '').strip()
            constraints = q.get('constraints', '').strip()
            sample_in = q.get('sample_input', '').strip()
            sample_out = q.get('sample_output', '').strip()

            c.execute("""INSERT INTO coding_questions
                         (test_no, class, section, subject, title, problem_statement,
                          allowed_languages, sample_input, sample_output, total_marks,
                          created_by, topic, difficulty, input_format, output_format,
                          constraints, ai_generated, approved)
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,0)""",
                      (test_no, class_, section, subject, title, statement,
                       allowed_langs, sample_in, sample_out, coding_marks,
                       uploaded_by, topic, coding_difficulty, inp_fmt, out_fmt, constraints))
            qid = c.lastrowid
            inserted_ids.append(qid)

            # Insert draft test cases
            tc_list = q.get('test_cases', [])
            if not tc_list and (sample_in or sample_out):
                tc_list = [{'input': sample_in, 'expected_output': sample_out, 'is_hidden': 0, 'marks_weight': 1.0}]
            
            for tc in tc_list:
                c.execute("""INSERT INTO coding_test_cases
                             (question_id, input, expected_output, is_hidden, marks_weight)
                             VALUES (?,?,?,?,?)""",
                          (qid, tc.get('input', ''), tc.get('expected_output', ''),
                           1 if tc.get('is_hidden') else 0, float(tc.get('marks_weight', 1.0))))

        # Register entry in test_papers repository
        c.execute("""INSERT INTO test_papers
                     (filename, class, section, subject, test_no, uploaded_by, uploader_type, is_active, question_count, paper_type)
                     VALUES (?,?,?,?,?,?,?,0,?,'Coding Test')""",
                  (f"{class_}{section}_{subject}_{test_no}_Coding", class_, section, subject, test_no, uploaded_by, uploader_type, len(inserted_ids)))

        conn.commit()
    except Exception as e:
        if conn:
            try:
                conn.rollback()
                conn.close()
            except Exception:
                pass
        return jsonify({'status': 'error', 'message': f'Server error saving coding questions: {str(e)}'}), 500

    conn.close()
    return jsonify({
        'status': 'success',
        'test_no': test_no,
        'count': len(inserted_ids),
        'redirect_url': f'/teacher/coding_review?test_no={test_no}',
        'message': f'Successfully generated {len(inserted_ids)} AI coding questions draft for Test {test_no}! Redirecting to Teacher Review screen...'
    })

@app.route('/teacher/coding_review')
def teacher_coding_review():
    """Teacher/Admin review, edit, and approval dashboard for draft coding questions."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return redirect(url_for('teacher_login'))

    test_no = request.args.get('test_no', '').strip()
    conn = get_db()
    c = conn.cursor()

    if test_no:
        c.execute("SELECT * FROM coding_questions WHERE test_no=? ORDER BY id ASC", (test_no,))
    else:
        c.execute("SELECT * FROM coding_questions WHERE approved=0 ORDER BY created_at DESC, id ASC")
    
    q_rows = [dict(r) for r in c.fetchall()]

    for q in q_rows:
        c.execute("SELECT * FROM coding_test_cases WHERE question_id=? ORDER BY id ASC", (q['id'],))
        q['test_cases'] = [dict(tc) for tc in c.fetchall()]

    conn.close()
    return render_template('coding_review.html', questions=q_rows, test_no=test_no, is_admin=bool(session.get('admin_logged_in')))

@app.route('/api/coding_questions/<int:qid>/update', methods=['POST'])
def api_update_coding_question(qid):
    """Update a draft coding question and its test cases."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403

    data = request.get_json(silent=True) or request.form
    title = data.get('title', '').strip()
    topic = data.get('topic', '').strip()
    difficulty = data.get('difficulty', 'medium').strip()
    statement = data.get('problem_statement', '').strip()
    inp_fmt = data.get('input_format', '').strip()
    out_fmt = data.get('output_format', '').strip()
    constraints = data.get('constraints', '').strip()
    sample_in = data.get('sample_input', '').strip()
    sample_out = data.get('sample_output', '').strip()
    total_marks = float(data.get('total_marks', 10))
    test_cases = data.get('test_cases', [])

    conn = get_db()
    c = conn.cursor()
    c.execute("""UPDATE coding_questions SET
                 title=?, topic=?, difficulty=?, problem_statement=?,
                 input_format=?, output_format=?, constraints=?,
                 sample_input=?, sample_output=?, total_marks=?
                 WHERE id=?""",
              (title, topic, difficulty, statement, inp_fmt, out_fmt, constraints, sample_in, sample_out, total_marks, qid))

    if test_cases and isinstance(test_cases, list):
        c.execute("DELETE FROM coding_test_cases WHERE question_id=?", (qid,))
        for tc in test_cases:
            c.execute("""INSERT INTO coding_test_cases
                         (question_id, input, expected_output, is_hidden, marks_weight)
                         VALUES (?,?,?,?,?)""",
                      (qid, tc.get('input', ''), tc.get('expected_output', ''),
                       1 if tc.get('is_hidden') else 0, float(tc.get('marks_weight', 1.0))))

    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Coding question updated successfully.'})

@app.route('/api/coding_questions/<int:qid>/approve', methods=['POST'])
def api_approve_coding_question(qid):
    """Approve a draft coding question, setting approved = 1 and marking test paper active."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403

    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE coding_questions SET approved=1 WHERE id=?", (qid,))
    c.execute("SELECT test_no FROM coding_questions WHERE id=?", (qid,))
    row = c.fetchone()
    if row and row['test_no']:
        c.execute("UPDATE test_papers SET is_active=1 WHERE test_no=?", (row['test_no'],))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': f'Coding Question #{qid} approved!'})

@app.route('/api/coding_questions/approve_all', methods=['POST'])
def api_approve_all_coding_questions():
    """Approve all draft coding questions for a specific test number or subject and activate paper."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403

    data = request.get_json(silent=True) or request.form
    test_no = data.get('test_no', '').strip()

    conn = get_db()
    c = conn.cursor()
    if test_no:
        c.execute("UPDATE coding_questions SET approved=1 WHERE test_no=?", (test_no,))
        c.execute("UPDATE test_papers SET is_active=1 WHERE test_no=?", (test_no,))
    else:
        c.execute("UPDATE coding_questions SET approved=1 WHERE approved=0")
        c.execute("UPDATE test_papers SET is_active=1 WHERE paper_type='Coding Test'")
    
    count = c.rowcount
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': f'Approved {count} coding questions and published test paper!'})

@app.route('/api/coding_questions/<int:qid>/delete', methods=['DELETE', 'POST'])
def api_delete_coding_question(qid):
    """Delete a coding question and its test cases."""
    if not session.get('admin_logged_in') and not session.get('teacher_logged_in'):
        return jsonify({'status': 'error', 'message': 'Not authorized'}), 403

    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM coding_questions WHERE id=?", (qid,))
    c.execute("DELETE FROM coding_test_cases WHERE question_id=?", (qid,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Coding question deleted.'})


# ---------------------------------------------------------------------------
# Student-Facing Coding Exam Portal Routes (Phase 3)
# ---------------------------------------------------------------------------

@app.route('/coding_exam_list')
@student_required
def coding_exam_list():
    """
    Render student question list view for a Coding Test session.
    """
    if not session.get('guidelines_accepted'):
        return redirect(url_for('guidelines'))

    student_id = session.get('student_id')
    conn = get_db()
    c = conn.cursor()

    active_sess = get_active_session_for_student(student_id=student_id, conn=conn)
    is_active = (active_sess is not None and active_sess.get('status') in ('running', 'paused'))
    if not is_active:
        c.execute("SELECT is_active FROM exam_control WHERE id=1")
        row = c.fetchone()
        if not row or not row['is_active']:
            conn.close()
            return redirect(url_for('waiting'))

    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return redirect(url_for('index'))

    if student['status'] == 'Submitted':
        conn.close()
        return redirect(url_for('submitted'))

    if student['status'] == 'Not Started':
        c.execute("UPDATE students SET status='In Progress', exam_started_at=CURRENT_TIMESTAMP WHERE student_id=?", (student_id,))
        conn.commit()

    # Fetch ordered questions for student
    c.execute("""SELECT cq.*, sq.shuffled_index 
                 FROM shuffled_questions sq
                 JOIN coding_questions cq ON sq.question_id = cq.id
                 WHERE sq.student_id = ? AND cq.approved = 1
                 ORDER BY sq.shuffled_index""", (student_id,))
    questions = [dict(r) for r in c.fetchall()]

    if not questions:
        clean_cls = (student['class'] or '').replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        c.execute("""SELECT * FROM coding_questions 
                     WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND test_no=? AND approved=1
                     ORDER BY id""",
                  (student['class'], f"%{clean_cls}%", student['subject'], student['test_no']))
        questions = [dict(r) for r in c.fetchall()]

    # Submission status mapping if table exists
    submissions_by_qid = {}
    try:
        c.execute("SELECT question_id, score, status FROM coding_submissions WHERE student_id=?", (student_id,))
        submissions_by_qid = {r['question_id']: dict(r) for r in c.fetchall()}
    except Exception:
        pass

    for q in questions:
        q['submission'] = submissions_by_qid.get(q['id'])

    conn.close()
    return render_template('coding_exam_list.html',
                           student=dict(student),
                           questions=questions,
                           active_sess=active_sess,
                           school_name=get_setting('school_name', 'RRB Group of Schools'),
                           logo_path=get_setting('logo_path', ''))


@app.route('/coding_editor/<int:question_id>')
@student_required
def coding_editor(question_id):
    """
    Render student split-screen Monaco code editor view for Phase 4.
    """
    if not session.get('guidelines_accepted'):
        return redirect(url_for('guidelines'))

    student_id = session.get('student_id')
    conn = get_db()
    c = conn.cursor()

    active_sess = get_active_session_for_student(student_id=student_id, conn=conn)
    is_active = (active_sess is not None and active_sess.get('status') in ('running', 'paused'))
    if not is_active:
        c.execute("SELECT is_active FROM exam_control WHERE id=1")
        row = c.fetchone()
        if not row or not row['is_active']:
            conn.close()
            return redirect(url_for('waiting'))

    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student or student['status'] == 'Submitted':
        conn.close()
        return redirect(url_for('submitted'))

    c.execute("SELECT * FROM coding_questions WHERE id=? AND approved=1", (question_id,))
    question = c.fetchone()
    if not question:
        conn.close()
        flash("Coding question not found or not approved.", "danger")
        return redirect(url_for('coding_exam_list'))

    # Strictly fetch VISIBLE test cases ONLY (is_hidden = 0). Hidden test cases are NEVER exposed.
    c.execute("SELECT input, expected_output FROM coding_test_cases WHERE question_id=? AND is_hidden=0 ORDER BY id", (question_id,))
    sample_cases = [dict(r) for r in c.fetchall()]

    conn.close()
    return render_template('coding_editor.html',
                           student=dict(student),
                           question=dict(question),
                           sample_cases=sample_cases,
                           active_sess=active_sess)


@app.route('/api/coding_drafts/save', methods=['POST'])
@student_required
def api_save_coding_draft():
    """Autosave or explicitly save code draft per student + question + language."""
    student_id = session.get('student_id')
    if not student_id:
        return jsonify({'status': 'error', 'message': 'Student session required'}), 401

    data = request.get_json(silent=True) or request.form
    question_id = data.get('question_id')
    language = (data.get('language') or 'python').strip().lower()
    code = data.get('code', '')

    if not question_id:
        return jsonify({'status': 'error', 'message': 'question_id is required'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("""INSERT INTO coding_drafts (student_id, question_id, language, code, updated_at)
                 VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                 ON CONFLICT(student_id, question_id, language) DO UPDATE SET
                     code = excluded.code,
                     updated_at = CURRENT_TIMESTAMP""",
              (student_id, int(question_id), language, code))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success', 'message': 'Draft saved successfully'})


@app.route('/api/coding_drafts/get/<int:question_id>', methods=['GET'])
@student_required
def api_get_coding_draft(question_id):
    """Retrieve saved code draft for a student + question + language."""
    student_id = session.get('student_id')
    if not student_id:
        return jsonify({'status': 'error', 'message': 'Student session required'}), 401

    language = request.args.get('lang', 'python').strip().lower()
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT code, updated_at FROM coding_drafts 
                 WHERE student_id=? AND question_id=? AND language=?""",
              (student_id, question_id, language))
    row = c.fetchone()
    conn.close()

    if row:
        return jsonify({'status': 'success', 'code': row['code'], 'updated_at': str(row['updated_at'])})
    return jsonify({'status': 'not_found', 'code': ''})


@app.route('/api/coding_run', methods=['POST'])
@student_required
def api_coding_run():
    """
    Execute student code snippet against custom/sample stdin input via Phase 1 runner.
    Does NOT compute grades or marks.
    """
    student_id = session.get('student_id')
    if not student_id:
        return jsonify({'status': 'error', 'message': 'Authentication required'}), 401

    data = request.get_json(silent=True) or request.form
    code = data.get('code', '')
    language = (data.get('language') or 'python').strip().lower()
    stdin_input = data.get('stdin_input', '')

    if not code:
        return jsonify({'status': 'error', 'message': 'Code snippet is empty'}), 400

    run_result = run_code(code, language, stdin_input, timeout_seconds=5)
    return jsonify(run_result)


def sync_student_coding_results(student_id, conn=None):
    """
    Recalculates total coding score for a student and updates the results reporting table.
    """
    close_db = False
    if conn is None:
        conn = get_db()
        close_db = True
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
        student = c.fetchone()
        if not student:
            return
            
        test_no = student['test_no'] or ''
        clean_cls = (student['class'] or '').replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
        
        c.execute("""SELECT id, total_marks FROM coding_questions
                     WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND (test_no=? OR topic=? OR test_no='') AND approved=1""",
                  (student['class'], f"%{clean_cls}%", student['subject'], test_no, test_no))
        approved_qs = c.fetchall()
        
        if not approved_qs:
            return
            
        total_test_marks = sum(float(q['total_marks'] or 10.0) for q in approved_qs) or 10.0
        total_q_count = len(approved_qs)
        approved_qids = [q['id'] for q in approved_qs]
        
        placeholders = ','.join(['?'] * len(approved_qids))
        c.execute(f"""SELECT SUM(marks_awarded) as total_score FROM coding_submissions
                       WHERE student_id=? AND question_id IN ({placeholders})""",
                   [student_id] + approved_qids)
        row_score = c.fetchone()
        score_sum = float(row_score['total_score']) if (row_score and row_score['total_score'] is not None) else 0.0
        percentage = round((score_sum / total_test_marks) * 100, 2)
        
        _save_or_update_result(c, student_id, student['name'], student['class'], student['section'],
                               student['subject'], score_sum, total_q_count, percentage, test_no)
        conn.commit()
    except Exception as e:
        app.logger.error(f"Error in sync_student_coding_results: {e}")
    finally:
        if close_db:
            conn.close()


@app.route('/api/coding_submit', methods=['POST'])
@student_required
def api_coding_submit():
    """
    Save student submission into coding_submissions and run Phase 5 evaluation engine.
    Computes test cases passed, marks awarded, per-case results breakdown,
    and aggregates total score into existing results table.
    """
    import json
    student_id = session.get('student_id')
    if not student_id:
        return jsonify({'status': 'error', 'message': 'Authentication required'}), 401

    data = request.get_json(silent=True) or request.form
    question_id = data.get('question_id')
    language = (data.get('language') or 'python').strip().lower()
    code = data.get('code', '')

    if not question_id or not code:
        return jsonify({'status': 'error', 'message': 'question_id and code are required'}), 400

    conn = get_db()
    c = conn.cursor()

    # 1. Fetch student & question info
    c.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    student = c.fetchone()
    if not student:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Student profile not found'}), 404

    c.execute("SELECT * FROM coding_questions WHERE id=? AND approved=1", (int(question_id),))
    question = c.fetchone()
    if not question:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Coding question not found or not approved'}), 404

    question_marks = float(question['total_marks'] or 10.0)

    # 2. Fetch ALL test cases (visible + hidden) for evaluation
    c.execute("""SELECT id, input, expected_output, is_hidden, marks_weight 
                 FROM coding_test_cases 
                 WHERE question_id=? 
                 ORDER BY id""", (int(question_id),))
    tc_rows = [dict(r) for r in c.fetchall()]

    if not tc_rows:
        sample_in = question['sample_input'] or ''
        sample_out = question['sample_output'] or ''
        tc_rows = [{'id': 1, 'input': sample_in, 'expected_output': sample_out, 'is_hidden': 0, 'marks_weight': 1.0}]

    # 3. Save to coding_drafts first
    c.execute("""INSERT INTO coding_drafts (student_id, question_id, language, code, updated_at)
                 VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                 ON CONFLICT(student_id, question_id, language) DO UPDATE SET
                     code = excluded.code,
                     updated_at = CURRENT_TIMESTAMP""",
              (student_id, int(question_id), language, code))

    # 4. Run Evaluation Engine
    eval_res = evaluate_submission(code, language, tc_rows, question_marks=question_marks, timeout_seconds=5)

    # 5. Sanitize results for DB storage & student UI (hide inputs/outputs for hidden test cases)
    sanitized_case_results = []
    for r in eval_res['results']:
        is_h = bool(r.get('is_hidden'))
        sanitized_case_results.append({
            'test_case_id': r.get('test_case_id'),
            'is_hidden': is_h,
            'passed': bool(r.get('passed')),
            'status': r.get('status'),
            'input': '[Hidden]' if is_h else r.get('input'),
            'expected_output': '[Hidden]' if is_h else r.get('expected_output'),
            'actual_output': '[Hidden]' if is_h else r.get('actual_output'),
            'execution_time_ms': round(float(r.get('execution_time_ms', 0)), 1),
            'stderr': '[Hidden]' if is_h else r.get('stderr')
        })

    json_results_str = json.dumps(sanitized_case_results)

    # 6. Save or update coding_submissions
    c.execute("SELECT id FROM coding_submissions WHERE student_id=? AND question_id=?", (student_id, int(question_id)))
    existing = c.fetchone()
    if existing:
        c.execute("""UPDATE coding_submissions 
                     SET code=?, language=?, test_cases_passed=?, total_test_cases=?, marks_awarded=?,
                         status=?, stdout_output=?, stderr_output=?, submitted_at=CURRENT_TIMESTAMP
                     WHERE id=?""",
                  (code, language, eval_res['passed_count'], eval_res['total_test_cases'], eval_res['marks_awarded'],
                   eval_res['status'], json_results_str, '', existing['id']))
    else:
        c.execute("""INSERT INTO coding_submissions 
                     (student_id, question_id, code, language, test_cases_passed, total_test_cases, marks_awarded, status, stdout_output, stderr_output, submitted_at)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)""",
                  (student_id, int(question_id), code, language, eval_res['passed_count'], eval_res['total_test_cases'], eval_res['marks_awarded'], eval_res['status'], json_results_str, ''))

    # 7. Aggregate Total Score across all coding questions in this test_no
    test_no = student['test_no'] or session.get('test_no', '')
    clean_cls = (student['class'] or '').replace('th','').replace('st','').replace('nd','').replace('rd','').strip()
    
    c.execute("""SELECT id, total_marks FROM coding_questions
                 WHERE (class=? OR class LIKE ?) AND LOWER(subject)=LOWER(?) AND (test_no=? OR topic=? OR test_no='') AND approved=1""",
              (student['class'], f"%{clean_cls}%", student['subject'], test_no, test_no))
    approved_qs = c.fetchall()
    
    total_test_marks = sum(float(q['total_marks'] or 10.0) for q in approved_qs) or 10.0
    total_q_count = len(approved_qs) or 1

    approved_qids = [q['id'] for q in approved_qs]
    score_sum = 0.0
    if approved_qids:
        placeholders = ','.join(['?'] * len(approved_qids))
        c.execute(f"""SELECT SUM(marks_awarded) as total_score FROM coding_submissions
                       WHERE student_id=? AND question_id IN ({placeholders})""",
                   [student_id] + approved_qids)
        row_score = c.fetchone()
        if row_score and row_score['total_score'] is not None:
            score_sum = float(row_score['total_score'])

    percentage = round((score_sum / total_test_marks) * 100, 2)

    # 8. Save/Update result in reporting pipeline table `results`
    _save_or_update_result(c, student_id, student['name'], student['class'], student['section'],
                           student['subject'], score_sum, total_q_count, percentage, test_no)

    conn.commit()
    conn.close()

    return jsonify({
        'status': 'success',
        'message': 'Evaluation complete!',
        'eval_status': eval_res['status'],
        'passed_count': eval_res['passed_count'],
        'total_test_cases': eval_res['total_test_cases'],
        'marks_awarded': eval_res['marks_awarded'],
        'question_marks': question_marks,
        'case_results': sanitized_case_results
    })



# ═══════════════════════════════════════════════════════════════════════════════
# QUESTION BANK MODULE ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

# ── AJAX: curriculum hierarchy helpers ────────────────────────────────────────

@app.route('/api/qb/subjects')
def qb_api_subjects():
    """Return subjects for board+class (used by upload form AJAX)."""
    if not (session.get('admin_logged_in') or session.get('teacher_logged_in')):
        return jsonify([])
    board = request.args.get('board', 'CBSE')
    class_num = request.args.get('class', '10')
    subjects = question_bank.get_curriculum_subjects(board, class_num)
    return jsonify(subjects)


@app.route('/api/qb/chapters')
def qb_api_chapters():
    """Return chapter suggestions for board+class+subject (autocomplete)."""
    if not (session.get('admin_logged_in') or session.get('teacher_logged_in')):
        return jsonify([])
    board = request.args.get('board', 'CBSE')
    class_num = request.args.get('class', '10')
    subject = request.args.get('subject', '')
    chapters = question_bank.get_chapter_suggestions(board, class_num, subject)
    return jsonify(chapters)


# ── MAIN BROWSE VIEW ──────────────────────────────────────────────────────────

@app.route('/question-bank')
def question_bank_browse():
    """Browsable Question Bank — Admin and Teacher access."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return redirect(url_for('teacher_login'))

    conn = get_db()
    c = conn.cursor()

    # Filter params
    f_board = request.args.get('board', '')
    f_class = request.args.get('class', '')
    f_subject = request.args.get('subject', '')
    f_chapter = request.args.get('chapter', '')
    f_type = request.args.get('type', '')

    # Build query for approved items
    where_clauses = ["approved = 1"]
    params = []
    if f_board:
        where_clauses.append("board = ?")
        params.append(f_board)
    if f_class:
        where_clauses.append("class = ?")
        params.append(f_class)
    if f_subject:
        where_clauses.append("subject = ?")
        params.append(f_subject)
    if f_chapter:
        where_clauses.append("chapter LIKE ?")
        params.append(f'%{f_chapter}%')
    if f_type:
        where_clauses.append("question_type = ?")
        params.append(f_type)

    where_sql = " AND ".join(where_clauses)
    c.execute(f"SELECT * FROM question_bank_items WHERE {where_sql} ORDER BY board, class, subject, chapter, question_type, id", params)
    items = c.fetchall()

    # Distinct filter options (only from approved items)
    c.execute("SELECT DISTINCT board FROM question_bank_items WHERE approved=1 ORDER BY board")
    boards = [r['board'] for r in c.fetchall()]
    c.execute("SELECT DISTINCT class FROM question_bank_items WHERE approved=1 ORDER BY CAST(class AS INTEGER)")
    classes = [r['class'] for r in c.fetchall()]
    c.execute("SELECT DISTINCT subject FROM question_bank_items WHERE approved=1 ORDER BY subject")
    subjects = [r['subject'] for r in c.fetchall()]
    c.execute("SELECT DISTINCT chapter FROM question_bank_items WHERE approved=1 ORDER BY chapter")
    chapters = [r['chapter'] for r in c.fetchall()]

    # Stats
    c.execute("SELECT COUNT(*) as cnt FROM question_bank_items WHERE approved=1")
    total_approved = c.fetchone()['cnt']
    c.execute("SELECT COUNT(*) as cnt FROM question_bank_items WHERE approved=0")
    total_pending = c.fetchone()['cnt']
    c.execute("SELECT COUNT(*) as cnt FROM question_bank_papers")
    total_papers = c.fetchone()['cnt']

    conn.close()

    return render_template('question_bank.html',
        items=items, boards=boards, classes=classes, subjects=subjects, chapters=chapters,
        f_board=f_board, f_class=f_class, f_subject=f_subject, f_chapter=f_chapter, f_type=f_type,
        total_approved=total_approved, total_pending=total_pending, total_papers=total_papers,
        is_admin=is_admin, is_teacher=is_teacher,
        all_boards=question_bank.get_curriculum_boards()
    )


# ── UPLOAD FORM & EXTRACTION JOB ─────────────────────────────────────────────

@app.route('/question-bank/upload', methods=['GET', 'POST'])
def question_bank_upload():
    """Upload a PDF to the Question Bank and trigger background extraction."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return redirect(url_for('teacher_login'))

    all_boards = question_bank.get_curriculum_boards()
    error = None
    success = None

    if request.method == 'POST':
        board = request.form.get('board', '').strip()
        class_num = request.form.get('class', '').strip()
        subject = request.form.get('subject', '').strip()
        chapter = request.form.get('chapter', '').strip()
        qtype = request.form.get('question_type', 'MCQ').strip()
        pdf_file = request.files.get('pdf_file')

        if not all([board, class_num, subject, chapter, qtype, pdf_file and pdf_file.filename]):
            error = "All fields are required, including the PDF file."
        elif not pdf_file.filename.lower().endswith('.pdf'):
            error = "Only PDF files are accepted."
        else:
            # Save PDF
            safe_fname = secure_filename(pdf_file.filename)
            pdf_path = os.path.join(QB_UPLOAD_FOLDER, safe_fname)
            pdf_file.save(pdf_path)

            # Insert paper record
            conn = get_db()
            c = conn.cursor()
            uploader = session.get('teacher_id') or 'admin'
            uploader_type = 'admin' if is_admin else 'teacher'
            c.execute("""INSERT INTO question_bank_papers
                         (uploaded_by, uploader_type, original_filename, board, class, subject, chapter, question_type, status)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'processing')""",
                      (str(uploader), uploader_type, safe_fname, board, class_num, subject, chapter, qtype))
            paper_id = c.lastrowid
            conn.commit()
            conn.close()

            log_audit_event(uploader_type, str(uploader), 'QB_UPLOAD',
                            'question_bank_papers', str(paper_id))

            # Launch background extraction thread
            def _extract_job(pdf_path, board, class_num, subject, chapter, qtype, paper_id):
                try:
                    items = question_bank.extract_questions_from_pdf(
                        pdf_path=os.path.abspath(pdf_path),
                        board=board, class_num=class_num, subject=subject,
                        chapter=chapter, question_type=qtype, paper_id=paper_id
                    )
                    conn2 = get_db()
                    c2 = conn2.cursor()
                    for item in items:
                        c2.execute("""INSERT INTO question_bank_items
                                      (paper_id, board, class, subject, chapter, question_type,
                                       raw_question_text, raw_options, raw_answer, image_path,
                                       ai_generated, approved, source_page)
                                      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, ?)""",
                                   (paper_id, item['board'], item['class'], item['subject'],
                                    item['chapter'], item['question_type'],
                                    item['raw_question_text'], item['raw_options'],
                                    item['raw_answer'], item['image_path'],
                                    item['source_page']))
                    c2.execute("UPDATE question_bank_papers SET status='processed' WHERE id=?", (paper_id,))
                    conn2.commit()
                    conn2.close()
                    app.logger.info(f"[QB] Extracted {len(items)} questions for paper_id={paper_id}")
                except Exception as e:
                    app.logger.error(f"[QB] Extraction failed paper_id={paper_id}: {e}")
                    try:
                        conn3 = get_db()
                        c3 = conn3.cursor()
                        c3.execute("UPDATE question_bank_papers SET status='failed' WHERE id=?", (paper_id,))
                        conn3.commit()
                        conn3.close()
                    except Exception:
                        pass

            t = threading.Thread(target=_extract_job, daemon=True,
                                 args=(pdf_path, board, class_num, subject, chapter, qtype, paper_id))
            t.start()

            return redirect(url_for('question_bank_review', paper_id=paper_id))

    return render_template('question_bank_upload.html',
        all_boards=all_boards, error=error, success=success,
        is_admin=is_admin, is_teacher=is_teacher
    )


# ── PER-QUESTION REVIEW SCREEN ────────────────────────────────────────────────

@app.route('/question-bank/review/<int:paper_id>')
def question_bank_review(paper_id):
    """Per-question review & tagging screen for a specific upload."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return redirect(url_for('teacher_login'))

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM question_bank_papers WHERE id=?", (paper_id,))
    paper = c.fetchone()
    if not paper:
        conn.close()
        return "Paper not found", 404

    c.execute("SELECT * FROM question_bank_items WHERE paper_id=? ORDER BY source_page, id", (paper_id,))
    items = c.fetchall()
    conn.close()

    all_boards = question_bank.get_curriculum_boards()
    status = paper['status']

    return render_template('question_bank_review.html',
        paper=paper, items=items, status=status,
        all_boards=all_boards, is_admin=is_admin, is_teacher=is_teacher
    )


@app.route('/question-bank/review/<int:paper_id>/status')
def question_bank_review_status(paper_id):
    """AJAX polling — returns extraction status and question count."""
    if not (session.get('admin_logged_in') or session.get('teacher_logged_in')):
        return jsonify({'status': 'error'}), 401
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT status FROM question_bank_papers WHERE id=?", (paper_id,))
    row = c.fetchone()
    c.execute("SELECT COUNT(*) as cnt FROM question_bank_items WHERE paper_id=?", (paper_id,))
    count_row = c.fetchone()
    conn.close()
    if not row:
        return jsonify({'status': 'not_found'})
    return jsonify({'status': row['status'], 'count': count_row['cnt']})


@app.route('/question-bank/review/save', methods=['POST'])
def question_bank_review_save():
    """Save per-question edits and bulk approve selected items."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return jsonify({'status': 'error', 'message': 'Unauthorized'}), 401

    data = request.get_json(force=True) or {}
    paper_id = data.get('paper_id')
    edits = data.get('edits', [])   # list of {id, chapter, question_type, approved}
    approve_all = data.get('approve_all', False)

    approver = str(session.get('teacher_id') or 'admin')
    now = datetime.datetime.now().isoformat()

    conn = get_db()
    c = conn.cursor()

    if approve_all:
        c.execute("""UPDATE question_bank_items SET approved=1, approved_by=?, approved_at=?
                     WHERE paper_id=?""", (approver, now, paper_id))
    else:
        for edit in edits:
            item_id = edit.get('id')
            if not item_id:
                continue
            chapter = edit.get('chapter', '').strip()
            qtype = edit.get('question_type', 'MCQ').strip()
            approved = 1 if edit.get('approved') else 0
            approved_by = approver if approved else ''
            approved_at = now if approved else None
            c.execute("""UPDATE question_bank_items
                         SET chapter=?, question_type=?, approved=?, approved_by=?, approved_at=?
                         WHERE id=?""",
                      (chapter, qtype, approved, approved_by, approved_at, item_id))

    conn.commit()
    c.execute("SELECT COUNT(*) as cnt FROM question_bank_items WHERE paper_id=? AND approved=1", (paper_id,))
    approved_count = c.fetchone()['cnt']
    conn.close()

    log_audit_event('admin' if is_admin else 'teacher', approver,
                    'QB_APPROVE', 'question_bank_items', str(paper_id))
    return jsonify({'status': 'ok', 'approved_count': approved_count})


@app.route('/question-bank/item/delete', methods=['POST'])
def question_bank_item_delete():
    """Delete a single question bank item."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return jsonify({'status': 'error'}), 401
    data = request.get_json(force=True) or {}
    item_id = data.get('id')
    if not item_id:
        return jsonify({'status': 'error', 'message': 'No item id'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM question_bank_items WHERE id=?", (item_id,))
    conn.commit()
    conn.close()
    return jsonify({'status': 'ok'})


# ── AI GENERATION FROM BANK REFERENCE ────────────────────────────────────────

@app.route('/question-bank/generate', methods=['GET', 'POST'])
def question_bank_generate():
    """AI generation of new questions using Question Bank items as reference context."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return redirect(url_for('teacher_login'))

    all_boards = question_bank.get_curriculum_boards()
    result_items = []
    error = None
    provider_used = None
    form_data = {}

    if request.method == 'POST':
        board = request.form.get('board', '').strip()
        class_num = request.form.get('class', '').strip()
        subject = request.form.get('subject', '').strip()
        chapter = request.form.get('chapter', '').strip()
        qtype = request.form.get('question_type', 'MCQ').strip()
        try:
            count = max(1, min(20, int(request.form.get('count', '5'))))
        except ValueError:
            count = 5

        form_data = {'board': board, 'class': class_num, 'subject': subject,
                     'chapter': chapter, 'question_type': qtype, 'count': count}

        # Fetch reference items from question_bank_items
        conn = get_db()
        c = conn.cursor()
        c.execute("""SELECT raw_question_text, raw_options FROM question_bank_items
                     WHERE approved=1 AND board=? AND class=? AND subject=?
                     AND chapter LIKE ? AND question_type=?
                     ORDER BY RANDOM() LIMIT 15""",
                  (board, class_num, subject, f'%{chapter}%', qtype))
        ref_items = [dict(r) for r in c.fetchall()]
        conn.close()

        # Build prompt and call AI
        prompt = question_bank.build_generation_prompt(
            ref_items, board, class_num, subject, chapter, qtype, count
        )
        res_dict, provider_used, err_msg = generate_ai_content(prompt, timeout=90)
        log_audit_event('admin' if is_admin else 'teacher',
                        str(session.get('teacher_id') or 'admin'),
                        f'QB_GENERATE via {provider_used or "none"}',
                        'question_bank_items', '')

        if not res_dict:
            error = err_msg or "AI generation failed. Please try again."
        else:
            raw_text = res_dict['candidates'][0]['content']['parts'][0]['text'].strip()
            result_items = question_bank.parse_ai_question_response(
                raw_text, qtype, board, class_num, subject, chapter
            )
            if not result_items:
                error = "AI returned a response but no valid questions could be parsed."

    return render_template('question_bank_generate.html',
        all_boards=all_boards, result_items=result_items,
        error=error, provider_used=provider_used, form_data=form_data,
        is_admin=is_admin, is_teacher=is_teacher
    )


@app.route('/question-bank/generate/save', methods=['POST'])
def question_bank_generate_save():
    """Save selected AI-generated questions into question_bank_items (pending review)."""
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return jsonify({'status': 'error'}), 401

    data = request.get_json(force=True) or {}
    questions = data.get('questions', [])
    if not questions:
        return jsonify({'status': 'error', 'message': 'No questions provided'}), 400

    conn = get_db()
    c = conn.cursor()
    saved = 0
    for q in questions:
        q_text = q.get('raw_question_text', '').strip()
        if not q_text:
            continue
        c.execute("""INSERT INTO question_bank_items
                     (paper_id, board, class, subject, chapter, question_type,
                      raw_question_text, raw_options, raw_answer, image_path,
                      ai_generated, approved, source_page)
                     VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?, '', 1, 0, 0)""",
                  (q.get('board', ''), q.get('class', ''), q.get('subject', ''),
                   q.get('chapter', ''), q.get('question_type', 'MCQ'),
                   q_text, q.get('raw_options', '[]'), q.get('raw_answer', '')))
        saved += 1
    conn.commit()
    conn.close()

    return jsonify({'status': 'ok', 'saved': saved,
                    'message': f'{saved} question(s) saved to Question Bank (pending review)'})


# ── WORD EXPORT WITH OMML EQUATIONS ──────────────────────────────────────────

@app.route('/question-bank/export/word', methods=['POST'])
def question_bank_export_word():
    """
    Export selected approved Question Bank items as a Word document
    with native OMML equations (not plain text or images).
    """
    is_admin = session.get('admin_logged_in')
    is_teacher = session.get('teacher_logged_in')
    if not (is_admin or is_teacher):
        return jsonify({'status': 'error'}), 401

    data = request.get_json(force=True) or {}
    item_ids = data.get('item_ids', [])
    title = data.get('title', 'Question Bank Export')
    if not item_ids:
        return jsonify({'status': 'error', 'message': 'No items selected'}), 400

    conn = get_db()
    c = conn.cursor()
    placeholders = ','.join(['?'] * len(item_ids))
    c.execute(f"SELECT * FROM question_bank_items WHERE id IN ({placeholders}) AND approved=1 ORDER BY chapter, question_type, id",
              item_ids)
    items = [dict(r) for r in c.fetchall()]
    conn.close()

    if not items:
        return jsonify({'status': 'error', 'message': 'No approved items found for given IDs'}), 404

    try:
        import docx as _docx
        from docx.shared import Pt, RGBColor, Inches, Mm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
    except ImportError:
        return jsonify({'status': 'error', 'message': 'python-docx not installed'}), 500

    doc = _docx.Document()
    # A4 setup
    for sec in doc.sections:
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)
        sec.top_margin = Mm(18)
        sec.bottom_margin = Mm(18)
        sec.left_margin = Mm(20)
        sec.right_margin = Mm(20)

    # Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(26, 35, 126)

    school = get_setting('school_name', 'RRB Group of Schools')
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"{school}  |  Question Bank Export")
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph()  # spacer

    # Group by chapter
    chapters_seen = {}
    for item in items:
        ch = item['chapter']
        if ch not in chapters_seen:
            chapters_seen[ch] = []
        chapters_seen[ch].append(item)

    q_global = 1
    for chapter_name, ch_items in chapters_seen.items():
        # Chapter header
        p_ch = doc.add_paragraph()
        p_ch.paragraph_format.space_before = Pt(14)
        p_ch.paragraph_format.keep_with_next = True
        r_ch = p_ch.add_run(f"Chapter: {chapter_name}")
        r_ch.bold = True
        r_ch.font.size = Pt(12)
        r_ch.font.color.rgb = RGBColor(26, 35, 126)

        for item in ch_items:
            q_text = item.get('raw_question_text', '')
            q_type = item.get('question_type', 'MCQ')

            # Question paragraph with math-aware runs
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(8)
            p_q.paragraph_format.space_after = Pt(2)

            # Question prefix run
            r_prefix = p_q.add_run(f"Q{q_global}. [{q_type}]  ")
            r_prefix.bold = True
            r_prefix.font.size = Pt(11)

            # Add question text with OMML math support
            question_bank.add_math_aware_runs(p_q, q_text, base_font_size=11)

            # Image (if any)
            img_path = item.get('image_path', '')
            if img_path:
                abs_img = os.path.join(os.path.dirname(os.path.abspath(__file__)), img_path)
                if os.path.exists(abs_img) and os.path.getsize(abs_img) > 0:
                    try:
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_picture(abs_img, width=Inches(3.0))
                    except Exception:
                        pass

            # MCQ options with math support
            raw_opts = item.get('raw_options', '[]')
            try:
                opts = json.loads(raw_opts) if raw_opts else []
            except Exception:
                opts = []
            for opt in opts:
                p_opt = doc.add_paragraph()
                p_opt.paragraph_format.left_indent = Mm(10)
                p_opt.paragraph_format.space_before = Pt(1)
                p_opt.paragraph_format.space_after = Pt(1)
                question_bank.add_math_aware_runs(p_opt, str(opt), base_font_size=10)

            q_global += 1

    # Footer
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(24)
    r_foot = p_foot.add_run(f"*** End of Question Bank Export ***\nGenerated by RRB CBT School System | {datetime.datetime.now().strftime('%d/%m/%Y')}")
    r_foot.italic = True
    r_foot.font.size = Pt(8.5)
    r_foot.font.color.rgb = RGBColor(150, 150, 150)

    # Save
    fname = f"QB_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    out_path = os.path.join(EXPORT_FOLDER, fname)
    doc.save(out_path)

    return send_file(out_path, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# ── PAPERS MANAGEMENT (Admin view of all uploads) ─────────────────────────────

@app.route('/question-bank/papers')
@admin_required
def question_bank_papers_list():
    """Admin view of all uploaded Question Bank papers."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM question_bank_papers ORDER BY upload_date DESC")
    papers = c.fetchall()
    conn.close()
    return render_template('question_bank.html', view='papers', papers=papers,
                           is_admin=True, is_teacher=False,
                           all_boards=question_bank.get_curriculum_boards())


@app.route('/question-bank/papers/<int:paper_id>/delete', methods=['POST'])
@admin_required
def question_bank_paper_delete(paper_id):
    """Admin: delete a paper and all its extracted questions."""
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM question_bank_items WHERE paper_id=?", (paper_id,))
    c.execute("DELETE FROM question_bank_papers WHERE id=?", (paper_id,))
    conn.commit()
    conn.close()
    log_audit_event('admin', 'admin', 'QB_DELETE_PAPER', 'question_bank_papers', str(paper_id))
    return jsonify({'status': 'ok'})


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() in ('true', '1', 'yes')
    app.run(host=host, port=port, debug=debug_mode)
