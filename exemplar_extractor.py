import os
import sys
import json
import base64
import logging
import re
import threading
import time
import urllib.request
import urllib.error

try:
    import fitz  # PyMuPDF
except ImportError:
    try:
        import pymupdf as fitz
    except ImportError:
        fitz = None

logger = logging.getLogger(__name__)
_index_lock = threading.Lock()


def _sanitize_name(name):
    """Sanitize subject/chapter/class strings for safe filesystem folder names."""
    if not name:
        return "general"
    clean = re.sub(r'[^a-zA-Z0-9_-]', '_', str(name).strip().lower())
    clean = re.sub(r'_+', '_', clean).strip('_')
    return clean or "general"


def _get_api_key():
    """Retrieve active Gemini / AI API key from apikey.env or .api_key file."""
    # 1. Check apikey.env
    script_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(script_dir, 'apikey.env')
    if os.path.exists(env_path):
        try:
            with open(env_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith('#') or not line:
                        continue
                    if 'GEMINI' in line.upper() and '=' in line:
                        k = line.split('=', 1)[1].strip().strip('"\'')
                        if k:
                            return k
                    elif '=' not in line and len(line) > 20:
                        return line
        except Exception:
            pass

    # 2. Check .api_key
    dot_key = os.path.join(script_dir, '.api_key')
    if os.path.exists(dot_key):
        try:
            k = open(dot_key, 'r', encoding='utf-8').read().strip()
            if k:
                return k
        except Exception:
            pass

    # 3. Check environment
    return os.environ.get('GEMINI_API_KEY', '').strip()


# ── STEP 1: EXTRACT IMAGES FROM PDF ───────────────────────────────────────────
def extract_images(pdf_path, subject, class_num, chapter_name, output_base="exemplar_bank/images"):
    """
    Extract all diagrams and illustrations from PDF pages using PyMuPDF.
    Automatically filters out tiny icons/borders/lines (width or height < 50px).
    Saves images in: exemplar_bank/images/{subject}/class{class_num}/{chapter_name}/
    
    Returns list of image metadata dicts:
    [{"file": rel_path, "abs_path": abs_path, "page": page_num, "index": img_idx, "width": width, "height": height}]
    """
    if not fitz:
        logger.error("PyMuPDF (fitz) is not installed. Cannot extract images from PDF.")
        return []

    if not os.path.exists(pdf_path):
        logger.error(f"PDF file not found at: {pdf_path}")
        return []

    script_dir = os.path.dirname(os.path.abspath(__file__))
    safe_subj = _sanitize_name(subject)
    safe_cls = _sanitize_name(f"class_{class_num}")
    safe_ch = _sanitize_name(chapter_name)[:40]

    if os.path.isabs(output_base):
        target_dir = os.path.join(output_base, safe_subj, safe_cls, safe_ch)
    else:
        target_dir = os.path.join(script_dir, output_base, safe_subj, safe_cls, safe_ch)

    os.makedirs(target_dir, exist_ok=True)

    extracted_images = []

    try:
        doc = fitz.open(pdf_path)
        for page_idx in range(len(doc)):
            page_num = page_idx + 1
            page = doc[page_idx]
            image_list = page.get_images(full=True)

            img_idx_on_page = 1
            for img_info in image_list:
                xref = img_info[0]
                try:
                    base_img = doc.extract_image(xref)
                    if not base_img:
                        continue

                    img_bytes = base_img.get("image")
                    img_ext = base_img.get("ext", "png").lower()
                    width = base_img.get("width", 0)
                    height = base_img.get("height", 0)

                    # Filter out tiny icons, horizontal rule lines, bullet symbols
                    if width < 50 or height < 50:
                        continue

                    img_filename = f"page_{page_num}_img_{img_idx_on_page}.{img_ext}"
                    abs_img_path = os.path.join(target_dir, img_filename)
                    rel_img_path = os.path.relpath(abs_img_path, script_dir).replace('\\', '/')

                    with open(abs_img_path, "wb") as f_out:
                        f_out.write(img_bytes)

                    extracted_images.append({
                        "file": rel_img_path,
                        "abs_path": abs_img_path,
                        "page": page_num,
                        "index": img_idx_on_page,
                        "width": width,
                        "height": height,
                        "xref": xref
                    })
                    img_idx_on_page += 1

                except Exception as e:
                    logger.debug(f"Error extracting image xref {xref} on page {page_num}: {e}")
                    continue

        doc.close()
        logger.info(f"Extracted {len(extracted_images)} valid diagrams/images from '{pdf_path}' to {target_dir}")

    except Exception as e:
        logger.error(f"Error opening or reading PDF '{pdf_path}': {e}")

    return extracted_images


# ── STEP 2: EXTRACT QUESTIONS WITH GEMINI VISION ──────────────────────────────
def _call_gemini_vision_page(image_bytes, prompt, api_key, timeout=90):
    """Send rendered PDF page image to Gemini Vision API and parse structured JSON."""
    if not api_key:
        return None, "No Gemini API key provided."

    candidate_models = [
        "gemini-3.6-flash",
        "gemini-3.5-flash-lite",
        "gemini-flash-latest",
        "gemini-2.5-pro",
        "gemini-flash-lite-latest",
        "gemini-2.5-flash",
        "gemini-1.5-flash"
    ]


    img_b64 = base64.b64encode(image_bytes).decode('utf-8')
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": "image/png",
                            "data": img_b64
                        }
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 4096
        }
    }

    payload_bytes = json.dumps(payload).encode('utf-8')
    last_err = ""

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
                    res_json = json.loads(resp_data)
                    cand = res_json.get('candidates', [])
                    if cand and 'content' in cand[0] and 'parts' in cand[0]['content']:
                        raw_text = cand[0]['content']['parts'][0]['text'].strip()
                        return raw_text, None
            except urllib.error.HTTPError as e:
                err_body = e.read().decode('utf-8')
                last_err = f"{model_name} HTTP {e.code}: {err_body[:200]}"
                if e.code == 429:
                    time.sleep(2)
                    continue
                break
            except Exception as e:
                last_err = str(e)
                break

    return None, last_err


def _fallback_text_extract_questions(doc, subject, class_num, chapter_name):
    """
    Offline fallback parser: extracts questions from PDF text directly using regex patterns
    when Gemini API key is missing or network call fails.
    """
    questions = []
    q_counter = 1

    for page_idx in range(len(doc)):
        page_num = page_idx + 1
        page = doc[page_idx]
        text = page.get_text("text")
        if not text.strip():
            continue

        # Search for question blocks (e.g. "1. ", "Q1.", "Question 1:", etc.)
        lines = text.split('\n')
        curr_q = None

        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            q_match = re.match(r'^(?:Q(?:uestion)?\.?\s*)?(\d+)[\.\)]\s*(.+)', line_str, re.IGNORECASE)
            if q_match:
                if curr_q:
                    questions.append(curr_q)
                q_num = q_match.group(1)
                q_body = q_match.group(2)
                curr_q = {
                    "question_number": q_num,
                    "question_text": q_body,
                    "question_type": "MCQ",
                    "options": [],
                    "has_image": False,
                    "image_description": "",
                    "page": page_num
                }
                q_counter += 1
            elif curr_q:
                # Check for MCQ option (a) / (b) / (c) / (d)
                opt_match = re.match(r'^\(([a-dABCD])\)\s*(.+)', line_str) or re.match(r'^([a-dABCD])[\.\)]\s*(.+)', line_str)
                if opt_match:
                    opt_letter = opt_match.group(1).lower()
                    opt_text = f"({opt_letter}) {opt_match.group(2)}"
                    curr_q["options"].append(opt_text)
                else:
                    curr_q["question_text"] += " " + line_str
                    if any(w in line_str.lower() for w in ['figure', 'diagram', 'given below', 'shown in fig', 'graph']):
                        curr_q["has_image"] = True

        if curr_q:
            questions.append(curr_q)

    # Classify question types
    for q in questions:
        if len(q.get("options", [])) >= 3:
            q["question_type"] = "MCQ"
        elif len(q.get("question_text", "").split()) > 40:
            q["question_type"] = "long"
        else:
            q["question_type"] = "short"

    return questions


def extract_questions_with_gemini(pdf_path, subject, class_num, chapter_name):
    """
    Render PDF pages to images and use Gemini Vision API to extract structured questions.
    Returns a consolidated list of question dicts across all pages.
    """
    if not fitz:
        logger.error("PyMuPDF (fitz) is not installed.")
        return []

    if not os.path.exists(pdf_path):
        logger.error(f"PDF file not found: {pdf_path}")
        return []

    all_questions = []
    api_key = _get_api_key()

    prompt = f"""You are an expert NCERT Exemplar question extraction assistant.
Extract all educational questions from this NCERT Exemplar page for Class {class_num} {subject}, Chapter: '{chapter_name}'.
For each question on this page, return a JSON object with:
- "question_number": string (e.g. "1", "2", "Q1", "15")
- "question_text": complete question text (preserving math formulas, chemical formulas, sub-questions)
- "question_type": "MCQ" | "short" | "long" | "case-study"
- "options": list of 4 choices formatted as ["(a) ...", "(b) ...", "(c) ...", "(d) ..."] if MCQ, else empty list []
- "has_image": boolean true if the question refers to or includes a diagram, graph, circuit, or figure on this page, else false
- "image_description": concise description of the diagram/graph/figure if has_image is true, else ""

OUTPUT FORMAT: Return ONLY a valid JSON array of objects like this:
[
  {{
    "question_number": "1",
    "question_text": "Which of the following is not a physical change?",
    "question_type": "MCQ",
    "options": ["(a) Boiling of water", "(b) Melting of ice", "(c) Dissolution of salt", "(d) Combustion of Liquefied Petroleum Gas"],
    "has_image": false,
    "image_description": ""
  }}
]
"""

    try:
        doc = fitz.open(pdf_path)

        # If no Gemini API key, use direct offline text parsing
        if not api_key:
            logger.warning("No Gemini API key found in apikey.env / .api_key. Using offline PyMuPDF text extractor.")
            fallback_qs = _fallback_text_extract_questions(doc, subject, class_num, chapter_name)
            doc.close()
            return fallback_qs

        for page_idx in range(len(doc)):
            page_num = page_idx + 1
            page = doc[page_idx]

            # Render page to high-res image (150 DPI)
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")

            raw_resp, err = _call_gemini_vision_page(img_bytes, prompt, api_key, timeout=60)
            if not raw_resp:
                logger.warning(f"Gemini Vision failed on page {page_num}: {err}. Falling back to text for this page.")
                continue

            # Clean JSON fences if present
            clean_text = raw_resp.strip()
            if "```" in clean_text:
                for part in clean_text.split("```"):
                    part = part.strip()
                    if part.startswith("json"):
                        part = part[4:].strip()
                    if part.startswith("[") and part.endswith("]"):
                        clean_text = part
                        break

            try:
                page_qs = json.loads(clean_text)
                if isinstance(page_qs, list):
                    for q in page_qs:
                        if isinstance(q, dict) and q.get("question_text"):
                            q["page"] = page_num
                            all_questions.append(q)
            except Exception as e:
                logger.warning(f"Error parsing JSON on page {page_num}: {e}")

        doc.close()

        # If Gemini didn't extract any questions (e.g. quota limit), fallback to offline text
        if not all_questions:
            doc = fitz.open(pdf_path)
            all_questions = _fallback_text_extract_questions(doc, subject, class_num, chapter_name)
            doc.close()

    except Exception as e:
        logger.error(f"Error processing PDF with Gemini Vision: {e}")

    logger.info(f"Extracted total {len(all_questions)} questions from '{pdf_path}'")
    return all_questions


# ── STEP 3: BUILD COMBINED INDEX & MATCH IMAGES ───────────────────────────────
def build_index(subject, class_num, chapter_name, questions_json, images_data, index_path="exemplar_bank_index.json"):
    """
    Store extracted questions and images in a combined JSON index file:
    {"subject": ..., "class": ..., "chapter": ..., "questions": [...], "images": [...]}
    Thread-safely appends / updates existing index without overwriting other chapters.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not os.path.isabs(index_path):
        target_index = os.path.join(script_dir, index_path)
    else:
        target_index = index_path

    # Group extracted images by page number
    images_by_page = {}
    for img in images_data:
        p = img.get("page", 1)
        images_by_page.setdefault(p, []).append(img)

    # Link images to questions on the same page
    for q in questions_json:
        p = q.get("page", 1)
        has_img = q.get("has_image", False)
        page_imgs = images_by_page.get(p, [])

        if has_img and page_imgs:
            # Assign first available unassigned image on this page, or best match
            matched_img = page_imgs[0]
            q["linked_image"] = matched_img.get("file")
            matched_img["linked_question_number"] = str(q.get("question_number", ""))
        elif page_imgs and not q.get("linked_image"):
            # Check if question text references diagram
            q_text = q.get("question_text", "").lower()
            if any(w in q_text for w in ["figure", "diagram", "graph", "fig.", "shown below", "given below"]):
                q["has_image"] = True
                matched_img = page_imgs[0]
                q["linked_image"] = matched_img.get("file")
                matched_img["linked_question_number"] = str(q.get("question_number", ""))

    chapter_entry = {
        "subject": str(subject).strip(),
        "class": str(class_num).strip().replace("Class", "").strip(),
        "chapter": str(chapter_name).strip(),
        "total_questions": len(questions_json),
        "total_images": len(images_data),
        "questions": questions_json,
        "images": images_data,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }

    with _index_lock:
        existing_index = []
        if os.path.exists(target_index):
            try:
                with open(target_index, 'r', encoding='utf-8') as f:
                    existing_index = json.load(f)
            except Exception as e:
                logger.error(f"Error reading existing exemplar index: {e}")
                existing_index = []

        if not isinstance(existing_index, list):
            existing_index = []

        # Replace existing entry for same subject + class + chapter, or append
        updated = False
        norm_subj = str(subject).strip().lower()
        norm_cls = str(class_num).strip().lower().replace("class", "").strip()
        norm_ch = str(chapter_name).strip().lower()

        for i, entry in enumerate(existing_index):
            if isinstance(entry, dict):
                e_subj = str(entry.get("subject", "")).strip().lower()
                e_cls = str(entry.get("class", "")).strip().lower().replace("class", "").strip()
                e_ch = str(entry.get("chapter", "")).strip().lower()
                if e_subj == norm_subj and e_cls == norm_cls and e_ch == norm_ch:
                    existing_index[i] = chapter_entry
                    updated = True
                    break

        if not updated:
            existing_index.append(chapter_entry)

        try:
            with open(target_index, 'w', encoding='utf-8') as f:
                json.dump(existing_index, f, indent=2, ensure_ascii=False)
            logger.info(f"Successfully saved Exemplar Bank index with {len(existing_index)} chapters at {target_index}")
        except Exception as e:
            logger.error(f"Failed to write exemplar index: {e}")

    return chapter_entry


# ── STEP 4: MASTER DRIVER FUNCTION ────────────────────────────────────────────
def process_exemplar_pdf(pdf_path, subject, class_num, chapter_name, index_path="exemplar_bank_index.json"):
    """
    Sequentially executes:
      1. extract_images() -> extracts and saves high-res diagrams in organized folders
      2. extract_questions_with_gemini() -> extracts questions with vision API
      3. build_index() -> builds and appends to persistent local exemplar bank index
    Prints summary and returns structured chapter entry.
    """
    print(f"\n=======================================================")
    print(f"[PROCESS] Processing NCERT Exemplar PDF")
    print(f"Subject: {subject} | Class: {class_num} | Chapter: {chapter_name}")
    print(f"File: {pdf_path}")
    print(f"=======================================================")

    # 1. Extract Images
    print("\n[Step 1/3] Extracting diagrams and figures from PDF...")
    images_data = extract_images(pdf_path, subject, class_num, chapter_name)
    print(f"[PASS] Extracted {len(images_data)} visual diagrams.")

    # 2. Extract Questions with Gemini Vision
    print("\n[Step 2/3] Extracting structured questions with Gemini Vision...")
    questions = extract_questions_with_gemini(pdf_path, subject, class_num, chapter_name)
    print(f"[PASS] Extracted {len(questions)} questions.")

    # 3. Build & Save Index
    print("\n[Step 3/3] Building combined tagged index...")
    chapter_entry = build_index(subject, class_num, chapter_name, questions, images_data, index_path=index_path)
    print(f"[PASS] Index updated successfully in {index_path}.")

    print(f"\n=======================================================")
    print(f"Exemplar Extraction Summary:")
    print(f"   - Subject:   {subject}")
    print(f"   - Class:     {class_num}")
    print(f"   - Chapter:   {chapter_name}")
    print(f"   - Questions: {len(questions)}")
    print(f"   - Images:    {len(images_data)}")
    print(f"=======================================================\n")


    return chapter_entry


# ── EXEMPLAR BANK QUERY & PAPER GENERATION HELPERS ────────────────────────────
def get_exemplar_chapters(subject=None, class_num=None, index_path="exemplar_bank_index.json"):
    """Retrieve list of available chapters in the exemplar bank."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target_index = os.path.join(script_dir, index_path) if not os.path.isabs(index_path) else index_path

    if not os.path.exists(target_index):
        return []

    try:
        with open(target_index, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return []

    if not isinstance(data, list):
        return []

    results = []
    norm_subj = str(subject).strip().lower() if subject else None
    norm_cls = str(class_num).strip().lower().replace("class", "").strip() if class_num else None

    for entry in data:
        if not isinstance(entry, dict):
            continue
        e_subj = str(entry.get("subject", "")).strip().lower()
        e_cls = str(entry.get("class", "")).strip().lower().replace("class", "").strip()

        if norm_subj and e_subj != norm_subj:
            continue
        if norm_cls and e_cls != norm_cls:
            continue

        results.append({
            "subject": entry.get("subject"),
            "class": entry.get("class"),
            "chapter": entry.get("chapter"),
            "total_questions": entry.get("total_questions", len(entry.get("questions", []))),
            "total_images": entry.get("total_images", len(entry.get("images", [])))
        })

    return results


def get_exemplar_questions(subject, class_num, chapter_name=None, question_type=None, count=None, index_path="exemplar_bank_index.json"):
    """
    Fetch questions from local Exemplar bank matching criteria.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target_index = os.path.join(script_dir, index_path) if not os.path.isabs(index_path) else index_path

    if not os.path.exists(target_index):
        return []

    try:
        with open(target_index, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return []

    if not isinstance(data, list):
        return []

    norm_subj = str(subject).strip().lower()
    norm_cls = str(class_num).strip().lower().replace("class", "").strip()
    norm_ch = str(chapter_name).strip().lower() if chapter_name else None
    norm_type = str(question_type).strip().lower() if question_type else None

    matching_qs = []

    for entry in data:
        if not isinstance(entry, dict):
            continue
        e_subj = str(entry.get("subject", "")).strip().lower()
        e_cls = str(entry.get("class", "")).strip().lower().replace("class", "").strip()
        e_ch = str(entry.get("chapter", "")).strip().lower()

        if e_subj != norm_subj or e_cls != norm_cls:
            continue
        if norm_ch and norm_ch not in e_ch and e_ch not in norm_ch:
            continue

        for q in entry.get("questions", []):
            if not isinstance(q, dict):
                continue
            q_type = str(q.get("question_type", "")).strip().lower()
            if norm_type and norm_type != q_type:
                continue

            q_copy = dict(q)
            q_copy["chapter"] = entry.get("chapter")
            q_copy["subject"] = entry.get("subject")
            q_copy["class"] = entry.get("class")
            matching_qs.append(q_copy)

    if count and count > 0:
        import random
        if len(matching_qs) > count:
            matching_qs = random.sample(matching_qs, count)

    return matching_qs


def generate_paper_from_exemplar(subject, class_num, chapter_name=None, mcq_count=10, short_count=5, long_count=3, index_path="exemplar_bank_index.json"):
    """
    Assemble a complete structured paper_data dict from local NCERT Exemplar bank questions.
    Resolves linked diagram file paths and placeholders.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    sections = []
    sec_letter = ord('A')

    # Section A: MCQs
    if mcq_count > 0:
        mcqs = get_exemplar_questions(subject, class_num, chapter_name=chapter_name, question_type="MCQ", count=mcq_count, index_path=index_path)
        if mcqs:
            q_list = []
            for i, q in enumerate(mcqs, start=1):
                img_path = None
                linked_img = q.get("linked_image")
                if linked_img:
                    cand = os.path.join(script_dir, linked_img) if not os.path.isabs(linked_img) else linked_img
                    if os.path.exists(cand):
                        img_path = cand

                placeholder = None
                if q.get("has_image") and not img_path:
                    placeholder = q.get("image_description") or f"Exemplar Diagram for Question {i}"

                q_list.append({
                    "number": i,
                    "question": q.get("question_text", ""),
                    "options": q.get("options", []),
                    "sub_questions": [],
                    "needs_image": bool(img_path or placeholder),
                    "image_path": img_path,
                    "image_placeholder": placeholder
                })

            sections.append({
                "section_label": f"Section {chr(sec_letter)}",
                "section_title": "Multiple Choice Questions (NCERT Exemplar)",
                "marks_per_question": 1,
                "instruction": "Select the correct option from the choices provided.",
                "questions": q_list
            })
            sec_letter += 1

    # Section B: Short Answer
    if short_count > 0:
        shorts = get_exemplar_questions(subject, class_num, chapter_name=chapter_name, question_type="short", count=short_count, index_path=index_path)
        if shorts:
            q_list = []
            for i, q in enumerate(shorts, start=1):
                img_path = None
                linked_img = q.get("linked_image")
                if linked_img:
                    cand = os.path.join(script_dir, linked_img) if not os.path.isabs(linked_img) else linked_img
                    if os.path.exists(cand):
                        img_path = cand

                placeholder = None
                if q.get("has_image") and not img_path:
                    placeholder = q.get("image_description") or f"Exemplar Diagram for Question {i}"

                q_list.append({
                    "number": i,
                    "question": q.get("question_text", ""),
                    "options": [],
                    "sub_questions": [],
                    "needs_image": bool(img_path or placeholder),
                    "image_path": img_path,
                    "image_placeholder": placeholder
                })

            sections.append({
                "section_label": f"Section {chr(sec_letter)}",
                "section_title": "Short Answer Questions (NCERT Exemplar)",
                "marks_per_question": 3,
                "instruction": "Answer all questions concisely with appropriate reasoning.",
                "questions": q_list
            })
            sec_letter += 1

    # Section C: Long Answer
    if long_count > 0:
        longs = get_exemplar_questions(subject, class_num, chapter_name=chapter_name, question_type="long", count=long_count, index_path=index_path)
        if longs:
            q_list = []
            for i, q in enumerate(longs, start=1):
                img_path = None
                linked_img = q.get("linked_image")
                if linked_img:
                    cand = os.path.join(script_dir, linked_img) if not os.path.isabs(linked_img) else linked_img
                    if os.path.exists(cand):
                        img_path = cand

                placeholder = None
                if q.get("has_image") and not img_path:
                    placeholder = q.get("image_description") or f"Exemplar Diagram for Question {i}"

                q_list.append({
                    "number": i,
                    "question": q.get("question_text", ""),
                    "options": [],
                    "sub_questions": [],
                    "needs_image": bool(img_path or placeholder),
                    "image_path": img_path,
                    "image_placeholder": placeholder
                })

            sections.append({
                "section_label": f"Section {chr(sec_letter)}",
                "section_title": "Long Answer / Case Study (NCERT Exemplar)",
                "marks_per_question": 5,
                "instruction": "Answer in detail with diagrams where required.",
                "questions": q_list
            })
            sec_letter += 1

    return {
        "class": str(class_num),
        "subject": str(subject),
        "chapter": chapter_name or "All Chapters",
        "sections": sections
    }
