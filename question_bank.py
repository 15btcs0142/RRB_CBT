"""
question_bank.py
================
Question Bank Module — PDF Extraction, AI Generation, and OMML Math Output.
"""

import os
import re
import json
import base64
import logging
import time
import threading
import shutil
import subprocess
import tempfile
import urllib.request
import urllib.error

try:
    import fitz
except ImportError:
    try:
        import pymupdf as fitz
    except ImportError:
        fitz = None

try:
    from lxml import etree as _lxml_etree
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False

logger = logging.getLogger(__name__)
QB_IMAGES_FOLDER = 'question_bank_images'


def _get_gemini_key():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(script_dir, 'apikey.env')
    if os.path.exists(env_path):
        try:
            with open(env_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' in line:
                        k, v = line.split('=', 1)
                        if 'GEMINI' in k.upper():
                            val = v.strip().strip('"\'')
                            if val:
                                return val
        except Exception:
            pass
    dot_key = os.path.join(script_dir, '.api_key')
    if os.path.exists(dot_key):
        try:
            k = open(dot_key, 'r', encoding='utf-8').read().strip()
            if k:
                return k
        except Exception:
            pass
    return os.environ.get('GEMINI_API_KEY', '').strip()


def _call_gemini_vision(image_bytes, prompt, api_key, timeout=90):
    if not api_key:
        return None, "No Gemini API key."
    candidate_models = [
        'gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-flash-latest',
        'gemini-2.5-flash', 'gemini-2.5-pro', 'gemini-1.5-flash'
    ]
    img_b64 = base64.b64encode(image_bytes).decode('utf-8')
    payload_bytes = json.dumps({
        "contents": [{"parts": [
            {"text": prompt},
            {"inline_data": {"mime_type": "image/png", "data": img_b64}}
        ]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 4096}
    }).encode('utf-8')
    last_err = ""
    for model in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        for attempt in range(2):
            try:
                req = urllib.request.Request(
                    url, data=payload_bytes,
                    headers={'Content-Type': 'application/json'}, method='POST'
                )
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    res = json.loads(resp.read().decode('utf-8'))
                    cands = res.get('candidates', [])
                    if cands and 'content' in cands[0]:
                        return cands[0]['content']['parts'][0]['text'].strip(), None
            except urllib.error.HTTPError as e:
                err_body = e.read().decode('utf-8')
                last_err = f"{model} HTTP {e.code}: {err_body[:200]}"
                if e.code == 429:
                    time.sleep(3)
                    continue
                break
            except Exception as e:
                last_err = str(e)
                break
    return None, last_err


def _regex_fallback_extraction(doc):
    questions = []
    for page_idx in range(len(doc)):
        page_num = page_idx + 1
        text = doc[page_idx].get_text("text")
        if not text.strip():
            continue
        lines = text.split('\n')
        curr_q = None
        for line in lines:
            ls = line.strip()
            if not ls:
                continue
            q_match = re.match(r'^(?:Q(?:uestion)?\.?\s*)?(\d+)[.)]\s*(.+)', ls, re.IGNORECASE)
            if q_match:
                if curr_q:
                    questions.append(curr_q)
                curr_q = {
                    "question_number": q_match.group(1),
                    "question_text": q_match.group(2),
                    "question_type": "MCQ",
                    "options": [],
                    "has_image": False,
                    "image_description": "",
                    "page": page_num
                }
            elif curr_q:
                opt_match = (re.match(r'^\(([a-dA-D])\)\s*(.+)', ls) or
                             re.match(r'^([a-dA-D])[.)]\s*(.+)', ls))
                if opt_match:
                    curr_q["options"].append(f"({opt_match.group(1).lower()}) {opt_match.group(2)}")
                else:
                    curr_q["question_text"] += " " + ls
                    if any(w in ls.lower() for w in ['figure', 'diagram', 'graph', 'circuit', 'shown']):
                        curr_q["has_image"] = True
        if curr_q:
            questions.append(curr_q)
    for q in questions:
        opts = q.get("options", [])
        if len(opts) >= 2:
            q["question_type"] = "MCQ"
        elif len(q.get("question_text", "").split()) > 40:
            q["question_type"] = "Descriptive"
        else:
            q["question_type"] = "Numerical"
    return questions


def _normalize_qtype(raw):
    raw = str(raw).strip().lower()
    if raw in ('mcq', 'multiple choice', 'objective'):
        return 'MCQ'
    if raw in ('numerical', 'numeric', 'integer', 'short'):
        return 'Numerical'
    if raw in ('descriptive', 'long', 'case-study', 'case study', 'essay'):
        return 'Descriptive'
    if raw in ('coding', 'programming', 'code'):
        return 'Coding'
    return 'MCQ'


def _extract_pdf_images(pdf_path, output_dir, script_dir):
    if not fitz:
        return []
    images = []
    try:
        doc = fitz.open(pdf_path)
        for page_idx in range(len(doc)):
            page_num = page_idx + 1
            page = doc[page_idx]
            img_idx = 1
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                try:
                    base_img = doc.extract_image(xref)
                    if not base_img:
                        continue
                    w = base_img.get("width", 0)
                    h = base_img.get("height", 0)
                    if w < 50 or h < 50:
                        continue
                    ext = base_img.get("ext", "png").lower()
                    fname = f"page_{page_num}_img_{img_idx}.{ext}"
                    abs_path = os.path.join(output_dir, fname)
                    with open(abs_path, "wb") as f:
                        f.write(base_img["image"])
                    rel_path = os.path.relpath(abs_path, script_dir).replace('\\', '/')
                    images.append({"page": page_num, "rel_path": rel_path, "abs_path": abs_path})
                    img_idx += 1
                except Exception:
                    continue
        doc.close()
    except Exception as e:
        logger.error(f"Image extraction error: {e}")
    return images


def extract_questions_from_pdf(pdf_path, board, class_num, subject, chapter, question_type,
                                 paper_id, output_image_base=None):
    """Extract questions from an arbitrary PDF using Gemini Vision + regex fallback."""
    if not fitz:
        logger.error("PyMuPDF not installed.")
        return []
    if not os.path.exists(pdf_path):
        logger.error(f"PDF not found: {pdf_path}")
        return []

    script_dir = os.path.dirname(os.path.abspath(__file__))
    img_base = output_image_base or os.path.join(script_dir, QB_IMAGES_FOLDER)
    paper_img_dir = os.path.join(img_base, f"paper_{paper_id}")
    os.makedirs(paper_img_dir, exist_ok=True)

    embedded_images = _extract_pdf_images(pdf_path, paper_img_dir, script_dir)
    api_key = _get_gemini_key()

    prompt = (
        f"You are an expert educational question extraction assistant.\n"
        f"Extract ALL questions from this exam paper page.\n"
        f"Board: {board}, Class: {class_num}, Subject: {subject}, Chapter/Topic: {chapter}\n\n"
        "For EACH question return a JSON object with:\n"
        '- "question_number": string\n'
        '- "question_text": complete text preserving math/chemical formulas\n'
        '- "question_type": "MCQ" | "Numerical" | "Descriptive" | "Coding"\n'
        '- "options": ["(a) ...", "(b) ...", "(c) ...", "(d) ..."] if MCQ else []\n'
        '- "raw_answer": correct answer if visible else ""\n'
        '- "has_image": true/false\n'
        '- "image_description": brief description or ""\n\n'
        "Return ONLY a valid JSON array, no commentary."
    )

    all_questions = []
    try:
        doc = fitz.open(pdf_path)
        if not api_key:
            logger.warning("No Gemini key — using regex extractor.")
            all_questions = _regex_fallback_extraction(doc)
            doc.close()
        else:
            for page_idx in range(len(doc)):
                page_num = page_idx + 1
                pix = doc[page_idx].get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")
                raw_resp, err = _call_gemini_vision(img_bytes, prompt, api_key, timeout=75)
                if not raw_resp:
                    logger.warning(f"Gemini Vision failed page {page_num}: {err}")
                    continue
                clean = raw_resp.strip()
                if "```" in clean:
                    for part in clean.split("```"):
                        part = part.strip()
                        if part.startswith("json"):
                            part = part[4:].strip()
                        if part.startswith("["):
                            clean = part
                            break
                try:
                    page_qs = json.loads(clean)
                    if isinstance(page_qs, list):
                        for q in page_qs:
                            if isinstance(q, dict) and q.get("question_text"):
                                q["page"] = page_num
                                all_questions.append(q)
                except Exception as e:
                    logger.warning(f"JSON parse error page {page_num}: {e}")
            if not all_questions:
                logger.warning("Gemini returned nothing — regex fallback.")
                all_questions = _regex_fallback_extraction(doc)
            doc.close()
    except Exception as e:
        logger.error(f"PDF processing error: {e}")

    images_by_page = {}
    for img in embedded_images:
        images_by_page.setdefault(img["page"], []).append(img)

    result_items = []
    for q in all_questions:
        page_num = q.get("page", 1)
        page_imgs = images_by_page.get(page_num, [])
        linked_img = page_imgs[0]["rel_path"] if (q.get("has_image") and page_imgs) else ""
        options_json = json.dumps(q.get("options", []))
        result_items.append({
            "board": board, "class": class_num, "subject": subject, "chapter": chapter,
            "question_type": _normalize_qtype(q.get("question_type", question_type)),
            "raw_question_text": q.get("question_text", "").strip(),
            "raw_options": options_json,
            "raw_answer": q.get("raw_answer", ""),
            "image_path": linked_img,
            "source_page": q.get("page", 1),
            "paper_id": paper_id, "ai_generated": 0, "approved": 0,
        })

    logger.info(f"Extracted {len(result_items)} questions from '{pdf_path}'")
    return result_items


def build_generation_prompt(reference_items, board, class_num, subject, chapter,
                             question_type, generate_count=5):
    """Build AI prompt for new question generation using bank items as reference context."""
    ref_lines = []
    for i, item in enumerate(reference_items[:15], start=1):
        q_text = item.get("raw_question_text", "")
        opts = []
        try:
            opts = json.loads(item.get("raw_options") or "[]")
        except Exception:
            pass
        line = f"{i}. {q_text}"
        if opts:
            line += " [Options: " + " | ".join(str(o) for o in opts) + "]"
        ref_lines.append(line)

    ref_block = "\n".join(ref_lines) if ref_lines else "(No reference items — generate from knowledge)"

    type_instructions = {
        "MCQ": 'Each question must have exactly 4 options labeled (a), (b), (c), (d) in the options array, and a "correct_answer" field.',
        "Numerical": 'Each question requires a single numerical answer. Include "correct_answer" with the numeric value. Leave "options": [].',
        "Descriptive": 'Each question is a descriptive/essay question. Leave "options": [] and "correct_answer": "".',
        "Coding": 'Each question is a programming problem. Leave "options": [] and "correct_answer": "".',
    }
    type_hint = type_instructions.get(question_type, type_instructions["MCQ"])

    return (
        "You are an expert educational question generator for school examinations.\n\n"
        f"REFERENCE QUESTIONS (context/style only — Board: {board}, Class: {class_num}, Subject: {subject}, Chapter: {chapter}):\n"
        f"{ref_block}\n\n"
        "RULES:\n"
        "1. Do NOT copy or paraphrase any reference question verbatim. Generate ORIGINAL questions.\n"
        "2. Match the difficulty level, concept depth, and style of the references.\n"
        "3. Cover the same topic but test different aspects or use different values/scenarios.\n"
        "4. Use LaTeX notation for ALL math: inline as $...$ and display as $$...$$.\n"
        f"5. {type_hint}\n\n"
        f"Generate exactly {generate_count} new {question_type} questions on \"{chapter}\" for Class {class_num} {subject} ({board}).\n\n"
        "Return ONLY a valid JSON array with NO extra text:\n"
        "[\n"
        "  {\n"
        f'    "question_text": "...",\n'
        f'    "question_type": "{question_type}",\n'
        '    "options": [],\n'
        '    "correct_answer": ""\n'
        "  }\n"
        "]"
    )


def parse_ai_question_response(raw_text, question_type, board, class_num, subject, chapter):
    """Parse AI-generated question JSON into question_bank_items-compatible dicts."""
    if not raw_text:
        return []
    clean = raw_text.strip()
    if "```" in clean:
        for part in clean.split("```"):
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("["):
                clean = part
                break
    match = re.search(r'\[.*\]', clean, re.DOTALL)
    if match:
        clean = match.group(0)
    try:
        items = json.loads(clean)
    except Exception:
        try:
            clean = re.sub(r',\s*([\]}])', r'\1', clean)
            items = json.loads(clean)
        except Exception as e:
            logger.error(f"Failed to parse AI response: {e}")
            return []
    result = []
    for item in items:
        if not isinstance(item, dict):
            continue
        q_text = item.get("question_text", "").strip()
        if not q_text:
            continue
        opts = item.get("options", [])
        opts_json = json.dumps(opts) if isinstance(opts, list) else "[]"
        result.append({
            "board": board, "class": class_num, "subject": subject, "chapter": chapter,
            "question_type": _normalize_qtype(item.get("question_type", question_type)),
            "raw_question_text": q_text, "raw_options": opts_json,
            "raw_answer": str(item.get("correct_answer", "")),
            "image_path": "", "source_page": 0,
            "paper_id": None, "ai_generated": 1, "approved": 0,
        })
    return result


# ── PANDOC OMML MATH ──────────────────────────────────────────────────────────

def _find_pandoc():
    p = shutil.which("pandoc")
    if p:
        return p
    for candidate in [
        r"C:\Program Files\Pandoc\pandoc.exe",
        r"C:\Program Files (x86)\Pandoc\pandoc.exe",
        os.path.expanduser(r"~\AppData\Local\Pandoc\pandoc.exe"),
    ]:
        if os.path.exists(candidate):
            return candidate
    return None


_PANDOC_PATH = _find_pandoc()
PANDOC_AVAILABLE = bool(_PANDOC_PATH)
_OMML_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def latex_segment_to_omml(latex_expr):
    """Convert LaTeX math string to OMML lxml Element via Pandoc. Returns None on failure."""
    if not PANDOC_AVAILABLE or not LXML_AVAILABLE:
        return None
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            md_file = os.path.join(tmpdir, "math.md")
            docx_file = os.path.join(tmpdir, "math.docx")
            with open(md_file, 'w', encoding='utf-8') as f:
                f.write(f"$${latex_expr}$$\n")
            result = subprocess.run(
                [_PANDOC_PATH, "-f", "markdown", "-t", "docx", "--quiet", "-o", docx_file, md_file],
                capture_output=True, timeout=15
            )
            if result.returncode != 0 or not os.path.exists(docx_file):
                return None
            import zipfile
            with zipfile.ZipFile(docx_file, 'r') as z:
                doc_xml = z.read("word/document.xml")
            root = _lxml_etree.fromstring(doc_xml)
            omath_elems = root.findall(f".//{{{_OMML_NS}}}oMath")
            return omath_elems[0] if omath_elems else None
    except Exception as e:
        logger.debug(f"OMML conversion error: {e}")
        return None


def _math_to_unicode(latex_expr):
    """Fallback: convert simple LaTeX to readable Unicode."""
    s = latex_expr
    for pat, rep in [
        (r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'(\1/\2)'),
        (r'\\sqrt\{([^{}]+)\}', r'sqrt(\1)'), (r'\\sqrt', 'sqrt'),
        (r'\\int', '∫'), (r'\\sum', '∑'), (r'\\prod', '∏'),
        (r'\\infty', '∞'), (r'\\pm', '±'), (r'\\times', '×'), (r'\\div', '÷'),
        (r'\\cdot', '·'), (r'\\approx', '≈'), (r'\\leq', '≤'), (r'\\geq', '≥'),
        (r'\\neq', '≠'), (r'\\alpha', 'α'), (r'\\beta', 'β'), (r'\\gamma', 'γ'),
        (r'\\delta', 'δ'), (r'\\Delta', 'Δ'), (r'\\theta', 'θ'), (r'\\lambda', 'λ'),
        (r'\\mu', 'μ'), (r'\\pi', 'π'), (r'\\sigma', 'σ'), (r'\\omega', 'ω'),
        (r'\\Omega', 'Ω'), (r'\\phi', 'φ'), (r'\\epsilon', 'ε'),
        (r'\\rightarrow|\\to', '→'), (r'\\leftarrow', '←'),
        (r'\\text\{([^{}]+)\}', r'\1'), (r'\\mathbf\{([^{}]+)\}', r'\1'),
        (r'\\\(|\\\)|\\\[|\\\]', ''), (r'\\[a-zA-Z]+', ''), (r'\{|\}', ''),
    ]:
        s = re.sub(pat, rep, s)
    return s.strip()


def split_text_math(text):
    """Split text into ('text', ...) and ('math', ...) segments."""
    segments = []
    pattern = re.compile(r'\$\$(.+?)\$\$|\$(.+?)\$', re.DOTALL)
    last_end = 0
    for match in pattern.finditer(text):
        start, end = match.span()
        if start > last_end:
            segments.append(('text', text[last_end:start]))
        latex_expr = match.group(1) if match.group(1) is not None else match.group(2)
        segments.append(('math', latex_expr.strip()))
        last_end = end
    if last_end < len(text):
        segments.append(('text', text[last_end:]))
    return segments


def add_math_aware_runs(paragraph, text, base_font_size=10):
    """
    Add runs to a python-docx paragraph with native OMML equations for $...$ math.
    Falls back to Unicode substitution if Pandoc/lxml unavailable.
    """
    if not text:
        return
    try:
        from docx.shared import Pt
    except ImportError:
        paragraph.add_run(str(text))
        return

    for seg_type, content in split_text_math(str(text)):
        if seg_type == 'text':
            if content:
                run = paragraph.add_run(content)
                run.font.size = Pt(base_font_size)
        else:
            omml_elem = latex_segment_to_omml(content)
            if omml_elem is not None:
                try:
                    omml_copy = _lxml_etree.fromstring(_lxml_etree.tostring(omml_elem))
                    paragraph._p.append(omml_copy)
                except Exception as e:
                    logger.debug(f"OMML splice failed: {e}")
                    run = paragraph.add_run(_math_to_unicode(content))
                    run.font.size = Pt(base_font_size)
            else:
                run = paragraph.add_run(_math_to_unicode(content))
                run.font.size = Pt(base_font_size)


# ── CURRICULUM HELPERS ────────────────────────────────────────────────────────

def get_chapter_suggestions(board, class_num, subject):
    """
    Return chapter suggestions with full subject alias matching,
    board normalization, and SQLite database fallback.
    """
    try:
        import curriculum_data
        chs = curriculum_data.get_chapters_and_topics(class_num, subject, board)
        suggestions = []
        if chs:
            for ch in chs:
                if isinstance(ch, dict) and ch.get("chapter"):
                    suggestions.append(ch["chapter"])
                elif isinstance(ch, str):
                    suggestions.append(ch)

        # Merge with existing chapters from database.db (questions and QB items)
        try:
            import sqlite3
            conn = sqlite3.connect('database.db')
            conn.row_factory = sqlite3.Row
            c = conn.cursor()
            c.execute("""
                SELECT DISTINCT chapter FROM questions 
                WHERE class = ? AND (LOWER(TRIM(subject)) = LOWER(TRIM(?)) OR LOWER(TRIM(subject)) LIKE ?) 
                AND chapter IS NOT NULL AND TRIM(chapter) != ''
            """, (str(class_num), str(subject), f"%{str(subject).lower()}%"))
            for r in c.fetchall():
                ch = r['chapter'].strip()
                if ch and ch not in suggestions:
                    suggestions.append(ch)

            c.execute("""
                SELECT DISTINCT chapter FROM question_bank_items 
                WHERE class = ? AND (LOWER(TRIM(subject)) = LOWER(TRIM(?)) OR LOWER(TRIM(subject)) LIKE ?)
                AND chapter IS NOT NULL AND TRIM(chapter) != ''
            """, (str(class_num), str(subject), f"%{str(subject).lower()}%"))
            for r in c.fetchall():
                ch = r['chapter'].strip()
                if ch and ch not in suggestions:
                    suggestions.append(ch)
            conn.close()
        except Exception:
            pass

        return suggestions
    except Exception:
        return []


def get_curriculum_boards():
    try:
        import curriculum_data
        return list(curriculum_data.CURRICULUM.keys())
    except Exception:
        return ["CBSE", "ICSE", "ICSE / ISC", "State / Other Board"]


def get_curriculum_classes(board):
    try:
        import curriculum_data
        return list(curriculum_data.CURRICULUM.get(board, {}).keys())
    except Exception:
        return [str(i) for i in range(1, 13)]


def get_curriculum_subjects(board, class_num):
    try:
        import curriculum_data
        return list(curriculum_data.CURRICULUM.get(board, {}).get(str(class_num), {}).keys())
    except Exception:
        return []
