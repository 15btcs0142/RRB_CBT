/**
 * RRB CBT — Multimodal Rendering Module  v2.0
 * =============================================
 * 1. Mathematics  — MathJax 3 (tex-svg) + mhchem extension
 * 2. Chemistry    — mhchem \ce{} equations + SmilesDrawer 2D structures
 * 3. Diagrams     — Imagen 3 via backend proxy + lightbox zoom
 */
'use strict';

// ── GEMINI SYSTEM INSTRUCTION ────────────────────────────────────────────────
const GEMINI_SYSTEM_INSTRUCTION = `
You are an expert academic question paper generator for CBSE/ICSE curriculum.
Output ONLY a valid JSON array — no markdown fences, no prose outside JSON.

SCHEMA (every field required):
{
  "question_type": "MCQ"|"Assertion-Reason"|"Very Short"|"Short"|"Long"|"Case Study",
  "content_type":  "math"|"chemistry"|"physics"|"biology"|"text",
  "question":      "<text with LaTeX or mhchem>",
  "option_a":      "<text or LaTeX, empty string for non-MCQ>",
  "option_b":      "<text or LaTeX, empty string for non-MCQ>",
  "option_c":      "<text or LaTeX, empty string for non-MCQ>",
  "option_d":      "<text or LaTeX, empty string for non-MCQ>",
  "correct_answer":"option_a"|"option_b"|"option_c"|"option_d"|"N/A",
  "smiles":        "<SMILES string or empty string>",
  "image_prompt":  "<Imagen prompt or empty string>",
  "marks":         1|2|3|4|5
}

══════════════════════════════════════════
MATHEMATICS LaTeX RULES (STRICTLY ENFORCED)
══════════════════════════════════════════
• Inline math  →  $...$
  "Find $\\frac{d}{dx}(x^2 + 3x)$ at $x = 2$"

• Block/display math  →  $$...$$
  "$$\\int_{0}^{\\pi} \\sin x\\,dx = 2$$"

• Matrices (pmatrix / bmatrix / vmatrix):
  "$A = \\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}$"

• Vectors: "$\\vec{F} = 3\\hat{i} - 4\\hat{j} + 5\\hat{k}$"

• Limits: "$\\lim_{x \\to 0} \\dfrac{\\sin x}{x} = 1$"

• Fractions: always \\frac{num}{den}

• Options that contain expressions must also use $...$:
  "option_a": "$x = \\frac{1}{2}$"

• NEVER write raw Unicode math (×, √, ∫) — use LaTeX commands.
• JSON strings: escape backslashes — write \\frac not \frac.

══════════════════════════════════════════
CHEMISTRY mhchem RULES (STRICTLY ENFORCED)
══════════════════════════════════════════
• Equations: use \\ce{} inside $ $
  "$\\ce{H2SO4 + 2NaOH -> Na2SO4 + 2H2O}$"
  "$\\ce{CH4 + 2O2 -> CO2 + 2H2O}$"
  "$\\ce{Fe^{2+} + 2e- -> Fe}$"
  "$\\ce{N2 + 3H2 <=> 2NH3}$"
  "$\\ce{CaCO3(s) -> CaO(s) + CO2(g)}$"

• Structural diagrams → set "smiles" field:
  Benzene="c1ccccc1", Ethanol="CCO",
  Acetic acid="CC(=O)O", Toluene="Cc1ccccc1",
  Aspirin="CC(=O)Oc1ccccc1C(=O)O",
  Glucose="OC[C@H]1OC(O)[C@H](O)[C@@H](O)[C@@H]1O"

BIOLOGY/PHYSICS DIAGRAMS:
• Set image_prompt with specific Imagen prompt:
  "Clean labeled diagram of human heart showing ventricles, aorta,
   pulmonary artery, vena cava, valves. Medical textbook style,
   white background, black labels, educational."

QUALITY:
• Assertion-Reason options must be:
  a) "Both A and R are true and R is correct explanation of A"
  b) "Both A and R are true but R is NOT correct explanation of A"
  c) "A is true but R is false"
  d) "A is false but R is true"
• Non-MCQ: set options="" and correct_answer="N/A"
• Mix difficulty 30% easy / 50% medium / 20% HOTS
`;

// ── MathJax STATE ────────────────────────────────────────────────────────────
let _mjReady   = false;
let _mjLoading = null;

function initMathJax() {
    if (_mjReady)   return Promise.resolve();
    if (_mjLoading) return _mjLoading;

    window.MathJax = {
        loader: { load: ['[tex]/mhchem','[tex]/boldsymbol','[tex]/cancel'] },
        tex: {
            packages:     {'[+]': ['mhchem','boldsymbol','cancel']},
            inlineMath:   [['$','$'],['\\(','\\)']],
            displayMath:  [['$$','$$'],['\\[','\\]']],
            processEscapes: true,
            tags: 'ams',
            macros: {
                ddx:  '\\dfrac{d}{dx}',
                dydx: '\\dfrac{dy}{dx}',
            }
        },
        svg: { fontCache:'global', scale:1.05, mtextInheritFont:true },
        options: {
            skipHtmlTags: ['script','noscript','style','textarea','pre','code'],
        },
        startup: {
            typeset: false,
            ready() {
                MathJax.startup.defaultReady();
                _mjReady = true;
            }
        }
    };

    _mjLoading = new Promise((resolve, reject) => {
        const s    = document.createElement('script');
        s.src      = 'https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js';
        s.async    = true;
        s.id       = 'MathJax-script';
        s.onload   = () => {
            const wait = (window.MathJax && MathJax.startup && MathJax.startup.promise)
                ? MathJax.startup.promise
                : new Promise(r => {
                    let n=0;
                    const t=setInterval(()=>{
                        n++;
                        if(_mjReady||(window.MathJax&&MathJax.typesetPromise)){
                            clearInterval(t); r();
                        } else if(n>80){ clearInterval(t); r(); }
                    },100);
                });
            wait.then(()=>{ _mjReady=true; resolve(); }).catch(resolve);
        };
        s.onerror  = ()=>reject(new Error('MathJax load failed'));
        document.head.appendChild(s);
    });

    return _mjLoading;
}

async function typeset(el) {
    if (!_mjReady || !window.MathJax?.typesetPromise) return;
    try { await MathJax.typesetPromise(el ? [el] : undefined); }
    catch(e){ console.warn('[MM] typeset:', e); }
}

// ── SmilesDrawer STATE ───────────────────────────────────────────────────────
let _sdReady   = false;
let _sdLoading = null;

function initSmilesDrawer() {
    if (_sdReady)   return Promise.resolve();
    if (_sdLoading) return _sdLoading;
    _sdLoading = new Promise((resolve, reject) => {
        const s    = document.createElement('script');
        s.src      = 'https://unpkg.com/smiles-drawer@2.1.7/dist/smiles-drawer.min.js';
        s.async    = true;
        s.onload   = ()=>{ _sdReady=true; resolve(); };
        s.onerror  = ()=>reject(new Error('SmilesDrawer load failed'));
        document.head.appendChild(s);
    });
    return _sdLoading;
}

// ── DETECTION ────────────────────────────────────────────────────────────────
const RE_MATH = /\$|\\\(|\\\[|\\(?:frac|int|sum|lim|vec|hat|sqrt|begin|alpha|beta|gamma|delta|theta|lambda|pi|sigma|omega|partial|nabla|infty|cdot|times)\b/;
const RE_CHEM = /\\ce\{/;

function hasMathOrChem(text) {
    if (!text) return false;
    return RE_MATH.test(text) || RE_CHEM.test(text);
}

function questionNeedsMath(q) {
    return ['question','option_a','option_b','option_c','option_d']
        .some(f => hasMathOrChem(q[f] || ''));
}

// ── CORE RENDER ──────────────────────────────────────────────────────────────
async function renderQuestionContent(el, q) {
    if (!el || !q) return;
    const { question='', smiles='', image_prompt='' } = q;

    // Always show text immediately
    el.innerHTML = '';
    const span = document.createElement('span');
    span.className = 'q-text';
    span.textContent = question;
    el.appendChild(span);

    // Math/Chem rendering
    if (questionNeedsMath(q)) {
        span.innerHTML = question;
        span.classList.add('mj-content');
        await initMathJax();
        await typeset(el);
    }

    // SMILES diagram
    if (smiles && smiles.trim()) {
        await _drawSmiles(el, smiles.trim());
    }

    // Imagen diagram
    if (image_prompt && image_prompt.trim()) {
        _fetchDiagram(el, image_prompt.trim());
    }
}

// ── SMILES DRAW ──────────────────────────────────────────────────────────────
async function _drawSmiles(container, smiles) {
    let loaded = true;
    try { await initSmilesDrawer(); } catch(e){ loaded=false; }
    if (!loaded || typeof SmilesDrawer === 'undefined') {
        _smilesText(container, smiles); return;
    }

    const wrap   = document.createElement('div');
    wrap.className = 'sd-wrap';
    const lbl   = document.createElement('div');
    lbl.className = 'sd-label';
    lbl.textContent = 'Molecular Structure';
    const id     = 'sd-' + Math.random().toString(36).slice(2,8);
    const canvas = document.createElement('canvas');
    canvas.id    = id;
    wrap.appendChild(lbl);
    wrap.appendChild(canvas);
    container.appendChild(wrap);

    try {
        SmilesDrawer.parse(smiles, tree => {
            const d = new SmilesDrawer.Drawer({
                width:280, height:220, bondThickness:1.4,
                themes:{light:{C:'#222',O:'#c0392b',N:'#2980b9',H:'#666',
                               S:'#e67e22',BACKGROUND:'#ffffff'}}
            });
            d.draw(tree, id, 'light', false);
            canvas.style.cursor = 'zoom-in';
            canvas.title = 'Click to zoom';
            canvas.onclick = () => openLightbox(canvas.toDataURL(), `Structure: ${smiles}`);
        }, err => { console.warn('[MM] SMILES:', err); _smilesText(container,smiles); wrap.remove(); });
    } catch(e){ _smilesText(container,smiles); wrap.remove(); }
}

function _smilesText(container, smiles) {
    const d = document.createElement('div');
    d.className = 'sd-fallback';
    d.textContent = 'Structure: ' + smiles;
    container.appendChild(d);
}

// ── IMAGEN FETCH ─────────────────────────────────────────────────────────────
function _fetchDiagram(container, prompt) {
    const wrap = document.createElement('div');
    wrap.className = 'diag-wrap';
    wrap.innerHTML = '<div class="diag-loading"><span class="spin"></span><span>Generating diagram…</span></div>';
    container.appendChild(wrap);

    fetch('/api/generate_diagram', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ prompt })
    })
    .then(r=>r.json())
    .then(data => {
        if (data.status==='success' && data.image) {
            wrap.innerHTML = `
                <img src="${data.image}" alt="${_esc(prompt.substring(0,80))}…"
                     title="Click to zoom" class="diag-img"
                     onclick="RRBMultimodal.openLightbox(this.src,this.alt)">
                <div class="diag-cap"><i class="fas fa-expand-alt"></i> Click to zoom</div>`;
        } else {
            wrap.innerHTML = `<div class="diag-err"><i class="fas fa-image"></i> Diagram unavailable</div>`;
        }
    })
    .catch(()=>{ wrap.innerHTML = '<div class="diag-err"><i class="fas fa-wifi"></i> Could not load diagram</div>'; });
}

function _esc(s){ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ── LIGHTBOX ─────────────────────────────────────────────────────────────────
function openLightbox(src, cap) {
    let lb = document.getElementById('rrbLB');
    if (!lb) {
        lb = document.createElement('div');
        lb.id = 'rrbLB';
        lb.innerHTML = `<div class="lb-inner" onclick="event.stopPropagation()">
            <button class="lb-x" onclick="document.getElementById('rrbLB').style.display='none'">&times;</button>
            <img id="lbImg" src="" alt=""><div id="lbCap"></div></div>`;
        lb.onclick = ()=>lb.style.display='none';
        document.body.appendChild(lb);
    }
    document.getElementById('lbImg').src = src;
    document.getElementById('lbImg').alt = cap||'';
    document.getElementById('lbCap').textContent = cap||'';
    lb.style.display = 'flex';
}

// ── BATCH RENDER ─────────────────────────────────────────────────────────────
async function renderAllQuestions(questions) {
    if (!questions?.length) return;
    const hasMath   = questions.some(questionNeedsMath);
    const hasSmiles = questions.some(q=>q.smiles&&q.smiles.trim());
    await Promise.all([
        hasMath   ? initMathJax()        : Promise.resolve(),
        hasSmiles ? initSmilesDrawer().catch(()=>{}) : Promise.resolve(),
    ]);
    for (let i=0; i<questions.length; i++) {
        const el = document.getElementById('q-'+i) || document.getElementById('question-'+(i+1));
        if (el) await renderQuestionContent(el, questions[i]);
    }
    if (hasMath) await typeset(null);
}

async function renderMathString(el, text) {
    if (!el) return;
    el.innerHTML = text;
    el.classList.add('mj-content');
    await initMathJax();
    await typeset(el);
}

// ── STYLES ───────────────────────────────────────────────────────────────────
(function _css(){
    if (document.getElementById('rrb-mm-css')) return;
    const s = document.createElement('style');
    s.id = 'rrb-mm-css';
    s.textContent = `
.q-text,.mj-content{line-height:1.9;font-size:1rem;display:block;}
mjx-container{overflow-x:auto;max-width:100%;margin:2px 0;}
mjx-container[display="true"]{display:block;text-align:center;margin:12px auto;overflow-x:auto;padding:4px 0;}
.sd-wrap{margin:12px 0 4px;text-align:center;}
.sd-label{font-size:.74rem;font-weight:700;color:#1b5e20;letter-spacing:.04em;text-transform:uppercase;margin-bottom:5px;}
.sd-wrap canvas{border-radius:10px;border:1.5px solid #c8e6c9;background:#fff;box-shadow:0 2px 8px rgba(0,0,0,.08);max-width:100%;}
.sd-fallback{font-family:monospace;font-size:.8rem;color:#555;background:#f5f5f5;padding:5px 10px;border-radius:6px;margin-top:6px;word-break:break-all;}
.diag-wrap{margin:10px 0;border:1.5px solid #e0e0e0;border-radius:12px;overflow:hidden;background:#fafafa;}
.diag-loading{display:flex;align-items:center;gap:10px;padding:16px 18px;font-size:.86rem;color:#667eea;}
.spin{width:18px;height:18px;border:3px solid #e0e0e0;border-top-color:#667eea;border-radius:50%;animation:mm-sp .7s linear infinite;flex-shrink:0;}
@keyframes mm-sp{100%{transform:rotate(360deg)}}
.diag-img{display:block;margin:8px auto;max-width:100%;max-height:280px;border-radius:8px;cursor:zoom-in;transition:opacity .2s;}
.diag-img:hover{opacity:.88;}
.diag-cap{text-align:center;font-size:.74rem;color:#888;padding:0 10px 8px;}
.diag-err{padding:14px 18px;color:#999;font-size:.82rem;display:flex;align-items:center;gap:8px;}
#rrbLB{position:fixed;inset:0;background:rgba(0,0,0,.88);z-index:9999;display:none;align-items:center;justify-content:center;cursor:pointer;}
.lb-inner{position:relative;max-width:92vw;max-height:92vh;cursor:default;display:flex;flex-direction:column;align-items:center;}
#lbImg{max-width:100%;max-height:80vh;border-radius:10px;box-shadow:0 20px 60px rgba(0,0,0,.6);}
#lbCap{color:rgba(255,255,255,.72);font-size:.8rem;margin-top:9px;text-align:center;max-width:70vw;}
.lb-x{position:absolute;top:-13px;right:-13px;width:30px;height:30px;border-radius:50%;background:#dc3545;color:#fff;border:none;font-size:1.2rem;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;}
`;
    document.head.appendChild(s);
})();

// ── PUBLIC API ───────────────────────────────────────────────────────────────
window.RRBMultimodal = {
    render:             renderQuestionContent,
    renderAll:          renderAllQuestions,
    renderMath:         renderMathString,
    initMathJax:        initMathJax,
    initSmilesDrawer:   initSmilesDrawer,
    typeset:            typeset,
    openLightbox:       openLightbox,
    needsMath:          hasMathOrChem,
    questionNeedsMath:  questionNeedsMath,
    SYSTEM_INSTRUCTION: GEMINI_SYSTEM_INSTRUCTION,
};
