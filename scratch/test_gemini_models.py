import os
import sys
import json
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import _get_gemini_key, call_gemini_generate_content

key = _get_gemini_key()
print("Key length:", len(key) if key else 0)

if key:
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={key}"
    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            models = [m.get('name', '').replace('models/', '') for m in data.get('models', []) if 'generateContent' in m.get('supportedGenerationMethods', [])]
            print("Available generateContent models:", models)
    except Exception as e:
        print("Error listing models:", e)

    print("\nTesting call_gemini_generate_content with prompt...")
    prompt = "Return a JSON array with 1 MCQ question about Class 10 ICSE Math Trigonometry."
    res, err = call_gemini_generate_content(prompt, key)
    print("Res:", bool(res))
    print("Err:", err)
    if res:
        txt = res['candidates'][0]['content']['parts'][0]['text']
        print("Returned text length:", len(txt))
        print("Text snippet:", repr(txt[:500]))
