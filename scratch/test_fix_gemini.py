import os
import sys
import json
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import _get_gemini_key, get_working_model, _build_ai_prompt, safe_json_loads

key = _get_gemini_key()

def fixed_call_gemini(prompt, api_key):
    detected = get_working_model(api_key)
    candidate_models = ['gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-3.5-flash-lite', 'gemini-flash-latest', 'gemini-2.5-pro']
    if detected and detected not in candidate_models:
        candidate_models.insert(0, detected)

    generation_config = {
        "temperature": 0.7,
        "maxOutputTokens": 8192,
        "response_mime_type": "application/json"
    }

    payload_bytes = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": generation_config
    }).encode('utf-8')

    last_error = ""
    for model_name in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        try:
            req = urllib.request.Request(
                url,
                data=payload_bytes,
                headers={'Content-Type': 'application/json'},
                method='POST'
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                resp_data = resp.read().decode('utf-8')
                return json.loads(resp_data), model_name, None
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8')
            last_error = f"{model_name} HTTP {e.code}: {err_body[:200]}"
            print(f"Model failed: {last_error}")

    return None, None, last_error

prompt = _build_ai_prompt("10", "A", "Mathematics", "Trigonometry", "1", "ICSE Class 10 Math", ["3 MCQ"])
res, model_used, err = fixed_call_gemini(prompt, key)
print("Result success?", bool(res))
print("Model used:", model_used)
print("Error:", err)

if res:
    raw_text = res['candidates'][0]['content']['parts'][0]['text']
    with open('scratch/raw_output.json', 'w', encoding='utf-8') as f:
        f.write(raw_text)
    print("Raw text saved to scratch/raw_output.json")
    parsed = safe_json_loads(raw_text)
    print("safe_json_loads success?", parsed is not None)
    if parsed:
        print("Questions count:", len(parsed))
        with open('scratch/parsed_output.json', 'w', encoding='utf-8') as f:
            json.dump(parsed, f, ensure_ascii=False, indent=2)
        print("Parsed result saved to scratch/parsed_output.json")
