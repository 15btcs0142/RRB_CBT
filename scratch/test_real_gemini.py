import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app import _get_gemini_key, _build_ai_prompt, generate_ai_content, safe_json_loads

key = _get_gemini_key()
print("Gemini API key found:", bool(key))

if key:
    prompt = _build_ai_prompt("10", "A", "Mathematics", "Trigonometry", "1", "ICSE Class 10 Math Test", ["5 MCQ"])
    print("Sending prompt to Gemini...")
    result, provider, err = generate_ai_content(prompt, timeout=60)
    print("Provider:", provider)
    print("Error:", err)
    if result:
        raw_text = result['candidates'][0]['content']['parts'][0]['text'].strip()
        print(f"Raw AI response length: {len(raw_text)}")
        print("Raw AI response snippet (first 1500 chars):")
        print(raw_text[:1500])
        
        parsed = safe_json_loads(raw_text)
        print("Parsed successfully?", parsed is not None)
        if parsed is None:
            print("FULL RAW TEXT THAT FAILED TO PARSE:")
            print(raw_text)
