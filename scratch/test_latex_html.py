import re
import os

def latex_to_html_math(text):
    if not text or not isinstance(text, str):
        return text

    # First, handle LaTeX brackets \(...\) and \[...\] and $...$
    # We strip outer delimiters or replace them
    
    # Common physics/math LaTeX command replacements
    replacements = [
        # Fractions: \frac{a}{b} -> <sup>a</sup>/<sub>b</sub> formatted inline block
        (r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'<span style="display:inline-block;vertical-align:-0.4em;text-align:center;margin:0 2px;"><span style="border-bottom:1px solid #000;display:block;padding:0 2px;font-size:0.9em;">\1</span><span style="display:block;padding:0 2px;font-size:0.9em;">\2</span></span>'),
        # Square root: \sqrt{x}
        (r'\\sqrt\{([^{}]+)\}', r'&radic;<span style="border-top:1px solid #000;padding-top:1px;margin-left:1px;">\1</span>'),
        # Vectors & bold: \vec{E}, \mathbf{B}
        (r'\\vec\{([^{}]+)\}', r'<b>\1&#8407;</b>'),
        (r'\\mathbf\{([^{}]+)\}', r'<b>\1</b>'),
        # Text block: \text{...}
        (r'\\text\{([^{}]+)\}', r'\1'),
        # Spacing: \, or \; or \quad
        (r'\\[,;]', ' '),
        (r'\\quad', '&nbsp;&nbsp;'),
        # Proportional to: \propto
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
    ]

    # Function to replace exponents: x^{y} or x^2
    def replace_exp(m):
        base, exp = m.group(1), m.group(2)
        return f"{base}<sup>{exp}</sup>"

    # Function to replace subscripts: x_{y} or x_0
    def replace_sub(m):
        base, sub = m.group(1), m.group(2)
        return f"{base}<sub>{sub}</sub>"

    result = text
    # Run fraction replacement 3 times to handle nested fractions
    for _ in range(3):
        for pattern, repl in replacements:
            result = re.sub(pattern, repl, result)

    # Exponents: e.g. 10^{-11} or E^2 or N^2
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)\^\{([^}]+)\}', replace_exp, result)
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)\^([0-9a-zA-Z])', replace_exp, result)

    # Subscripts: e.g. v_d or V_{rms} or R_1
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)_\{([^}]+)\}', replace_sub, result)
    result = re.sub(r'([a-zA-Z0-9&;\>\}\)]+)_([0-9a-zA-Z])', replace_sub, result)

    # Clean up LaTeX delimiters \(...\), \[...\], $$...$$, $...$
    result = re.sub(r'\\\[|\\\]|\\\(|\\\)', '', result)
    result = re.sub(r'\$\$|\$', '', result)

    return result

# Test on the actual Class 12 Physics paper content
file_path = r'c:\Users\atuls\Desktop\RRB_CBT\RRB_cbt-v1.10\RRB_v110\paper\Class_12\Section_A\12_A_Physics_Pre-Board_GAURAV_SHUKLA.doc'
if os.path.exists(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    converted = latex_to_html_math(content)
    print("=== CONVERTED FILE PREVIEW (First 2500 chars) ===")
    print(converted[:2500])
