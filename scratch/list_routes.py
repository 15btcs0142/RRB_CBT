import os
import sys
import types

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
os.chdir(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    import pdfkit
except ImportError:
    sys.modules['pdfkit'] = types.ModuleType('pdfkit')

import app

with open('scratch/all_routes.txt', 'w', encoding='utf-8') as f:
    for rule in sorted(app.app.url_map.iter_rules(), key=lambda r: r.rule):
        methods = sorted(list(rule.methods - {'HEAD', 'OPTIONS'}))
        line = f"{rule.rule:<50} -> {rule.endpoint:<35} {methods}\n"
        f.write(line)
        print(line, end='')

print("\nSaved all routes to scratch/all_routes.txt")
