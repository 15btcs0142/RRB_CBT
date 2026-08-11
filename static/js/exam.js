// ========================================
// COMPREHENSIVE COPY/PASTE PREVENTION
// ========================================

// Disable right-click context menu
document.addEventListener('contextmenu', e => e.preventDefault());

// Disable text selection
document.addEventListener('selectstart', e => e.preventDefault());
document.addEventListener('mousedown', e => {
    if (e.detail > 1) {  // Prevent double/triple click selection
        e.preventDefault();
    }
});

// Disable all copy/cut/paste keyboard shortcuts
document.addEventListener('keydown', function(e) {
    // Prevent F12, Ctrl+Shift+I, Ctrl+U (dev tools)
    if (e.key === 'F12' || 
        (e.ctrlKey && e.shiftKey && e.key === 'I') || 
        (e.ctrlKey && e.key === 'U')) {
        e.preventDefault();
        return false;
    }
    
    // Prevent Ctrl+C, Ctrl+X, Ctrl+V, Ctrl+A (copy/cut/paste/select all)
    if (e.ctrlKey && (e.key === 'c' || e.key === 'C' ||
                      e.key === 'x' || e.key === 'X' ||
                      e.key === 'v' || e.key === 'V' ||
                      e.key === 'a' || e.key === 'A')) {
        e.preventDefault();
        return false;
    }
    
    // Prevent Cmd+C, Cmd+X, Cmd+V on Mac
    if (e.metaKey && (e.key === 'c' || e.key === 'C' ||
                      e.key === 'x' || e.key === 'X' ||
                      e.key === 'v' || e.key === 'V')) {
        e.preventDefault();
        return false;
    }
});

// Disable copy/cut/paste events
document.addEventListener('copy', e => {
    e.preventDefault();
    return false;
});
document.addEventListener('cut', e => {
    e.preventDefault();
    return false;
});
document.addEventListener('paste', e => {
    e.preventDefault();
    return false;
});

// Disable drag events
document.addEventListener('dragstart', e => e.preventDefault());

// Add CSS to prevent text selection
const style = document.createElement('style');
style.textContent = `
    * {
        -webkit-user-select: none !important;
        -moz-user-select: none !important;
        -ms-user-select: none !important;
        user-select: none !important;
    }
    input, textarea {
        -webkit-user-select: text !important;
        -moz-user-select: text !important;
        -ms-user-select: text !important;
        user-select: text !important;
    }
`;
document.head.appendChild(style);

// ========================================
// END COPY/PASTE PREVENTION
// ========================================


// Detect mobile device
function isMobileDevice() {
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) 
           || ('ontouchstart' in window) 
           || (navigator.maxTouchPoints > 0);
}

const isMobile = isMobileDevice();
let fullscreenEnabled = false;

// ── FULLSCREEN MANAGER ─────────────────────────────────────────────────────
// Bug #1 Fix: accidental focus loss (OS notifications, alt-tab, etc.) must NOT
// immediately submit. Instead: show a warning overlay and allow the student to
// re-enter fullscreen within a grace period. Only submit after 3 confirmed exits
// with no recovery, protecting students from spurious notification clicks.

if (!isMobile) {
    let _fsExitCount    = 0;          // consecutive exits without recovery
    let _fsGraceTimer   = null;       // timeout for grace period
    let _fsWarningShown = false;
    const FS_GRACE_MS   = 8000;       // 8-second grace window per exit
    const FS_MAX_EXITS  = 3;          // submit only after this many unrecovered exits

    function requestFullscreen() {
        const elem = document.documentElement;
        if      (elem.requestFullscreen)       elem.requestFullscreen();
        else if (elem.webkitRequestFullscreen) elem.webkitRequestFullscreen();
        else if (elem.msRequestFullscreen)     elem.msRequestFullscreen();
    }

    function isInFullscreen() {
        return !!(document.fullscreenElement   ||
                  document.webkitFullscreenElement ||
                  document.mozFullScreenElement    ||
                  document.msFullscreenElement);
    }

    function enterFullscreen() {
        requestFullscreen();
        setTimeout(() => {
            if (!isInFullscreen()) {
                showFullscreenPrompt(true);  // initial prompt — not an exit warning
            } else {
                fullscreenEnabled = true;
            }
        }, 500);
    }

    function showFullscreenPrompt(isInitial) {
        if (document.getElementById('fullscreenPrompt')) return;
        const div   = document.createElement('div');
        div.id      = 'fullscreenPrompt';
        div.style.cssText = `
            position:fixed;top:0;left:0;width:100%;height:100%;
            background:rgba(0,0,0,0.88);z-index:99999;
            display:flex;align-items:center;justify-content:center;`;
        div.innerHTML = `
            <div style="background:white;padding:36px 40px;border-radius:18px;
                        box-shadow:0 20px 60px rgba(0,0,0,0.5);text-align:center;max-width:420px;">
                <div style="font-size:2.5rem;margin-bottom:12px;">⚠️</div>
                <h3 style="color:#b71c1c;margin-bottom:10px;">
                    ${isInitial ? 'Fullscreen Required' : 'Fullscreen Exited'}
                </h3>
                <p style="color:#555;margin-bottom:20px;line-height:1.6;">
                    ${isInitial
                        ? 'This exam must be taken in fullscreen mode to continue.'
                        : 'You exited fullscreen. Please return to fullscreen immediately.<br><small style="color:#e53935;">Repeated exits will automatically submit your exam.</small>'}
                </p>
                <button id="fsBtnRe" style="background:#1b5e20;color:white;border:none;
                        padding:12px 28px;border-radius:30px;font-size:1rem;cursor:pointer;
                        display:inline-flex;align-items:center;gap:8px;">
                    ⛶ Return to Fullscreen
                </button>
            </div>`;
        document.body.appendChild(div);
        document.getElementById('fsBtnRe').addEventListener('click', () => {
            requestFullscreen();
            setTimeout(() => {
                if (isInFullscreen()) {
                    fullscreenEnabled = true;
                    _fsExitCount = 0;
                    clearTimeout(_fsGraceTimer);
                    _fsWarningShown = false;
                    const p = document.getElementById('fullscreenPrompt');
                    if (p) p.remove();
                }
            }, 400);
        });
    }

    function exitHandler() {
        if (isInFullscreen()) {
            // Entering fullscreen — clear any pending grace timer
            clearTimeout(_fsGraceTimer);
            _fsExitCount = 0;
            _fsWarningShown = false;
            const p = document.getElementById('fullscreenPrompt');
            if (p) p.remove();
            return;
        }

        // Student exited fullscreen
        if (sessionStorage.getItem('examSubmitted')) return;

        _fsExitCount++;

        if (_fsExitCount >= FS_MAX_EXITS) {
            // Too many exits — submit now
            const p = document.getElementById('fullscreenPrompt');
            if (p) p.remove();
            alert('You exited fullscreen too many times. The exam has been submitted.');
            silentSubmit();
            return;
        }

        // Show warning overlay and start grace timer
        showFullscreenPrompt(false);

        clearTimeout(_fsGraceTimer);
        _fsGraceTimer = setTimeout(() => {
            // Grace period expired without recovery — still in exit
            if (!isInFullscreen() && !sessionStorage.getItem('examSubmitted')) {
                // Give one more chance silently; increment on next exitHandler call
            }
        }, FS_GRACE_MS);
    }

    enterFullscreen();

    document.addEventListener('fullscreenchange',       exitHandler);
    document.addEventListener('webkitfullscreenchange', exitHandler);
    document.addEventListener('mozfullscreenchange',    exitHandler);
    document.addEventListener('MSFullscreenChange',     exitHandler);

    // visibilitychange: page hidden (e.g. notification click) — do NOT submit
    document.addEventListener('visibilitychange', () => {
        if (document.hidden) {
            // page lost focus — only increment exit count if already out of FS
            if (!isInFullscreen() && !sessionStorage.getItem('examSubmitted')) {
                // handled by exitHandler when FS state actually changes
            }
        }
    });
} else {
    console.log('Mobile device detected – fullscreen not enforced.');
}

// Global variables
let questions = [];
let currentIndex = 0;
let responses = {};
let markedForReview = new Set();
let visited = new Set();
let timerInterval;
let examDuration = 0;

// Mobile palette elements
let desktopGrid, mobileGrid;
let paletteToggleBtn, paletteOverlay, mobilePalette, closePaletteBtn;

function initMobilePalette() {
    paletteToggleBtn = document.getElementById('paletteToggleBtn');
    paletteOverlay = document.getElementById('paletteOverlay');
    mobilePalette = document.getElementById('mobilePalette');
    closePaletteBtn = document.getElementById('closePaletteBtn');

    if (paletteToggleBtn) {
        paletteToggleBtn.addEventListener('click', function() {
            paletteOverlay.classList.add('active');
            mobilePalette.classList.add('active');
        });
    }
    if (closePaletteBtn) {
        closePaletteBtn.addEventListener('click', function() {
            paletteOverlay.classList.remove('active');
            mobilePalette.classList.remove('active');
        });
    }
    if (paletteOverlay) {
        paletteOverlay.addEventListener('click', function() {
            paletteOverlay.classList.remove('active');
            mobilePalette.classList.remove('active');
        });
    }
}

// Fetch questions and initialize
fetch('/get_questions')
    .then(res => res.json())
    .then(data => {
        questions = data;
        window.questions = questions;   // expose for combined tab JS
        if (questions.length === 0) {
            alert('No questions available for this subject/class.');
            window.location.href = '/';
            return;
        }
        questions.forEach((q, idx) => {
            if (q.selected) {
                responses[idx] = q.selected;
            }
        });
        const totalHeader = document.getElementById('questionTotalHeader');
        const totalPanel = document.getElementById('questionTotalPanel');
        if (totalHeader) totalHeader.textContent = questions.length;
        if (totalPanel) totalPanel.textContent = questions.length;
        initMobilePalette();
        renderPalette();
        loadQuestion(0);
        startTimer();
        questions.forEach((_, i) => updatePaletteButton(i));
        // Trigger combined-test tab init if applicable
        if (typeof window.initExam === 'function') window.initExam();
    });

function startTimer() {
    fetch('/get_exam_time')
        .then(res => res.json())
        .then(data => {
            // BUG-002 FIX: if server says remaining=0 AND student is 'Submitted', auto-submit
            if (data.force_submitted) {
                sessionStorage.setItem('examSubmitted', 'true');
                clearInterval(timerInterval);
                clearInterval(statusPollInterval);
                window.location.href = '/submitted';
                return;
            }
            examDuration = data.remaining;
            if (examDuration <= 0) {
                alert('Time is up! Your exam will be submitted now.');
                silentSubmit();
                return;
            }
            updateTimerDisplay();
            timerInterval = setInterval(() => {
                examDuration--;
                updateTimerDisplay();
                if (examDuration <= 0) {
                    clearInterval(timerInterval);
                    alert('Time is up! Your exam will be submitted now.');
                    silentSubmit();
                }
            }, 1000);
        });
}

// BUG-002 FIX: Poll every 5s to detect admin-forced submission
let statusPollInterval = setInterval(() => {
    if (sessionStorage.getItem('examSubmitted')) {
        clearInterval(statusPollInterval);
        return;
    }
    fetch('/check_exam_status')
        .then(r => r.json())
        .then(data => {
            if (data.force_submitted) {
                clearInterval(timerInterval);
                clearInterval(statusPollInterval);
                sessionStorage.setItem('examSubmitted', 'true');
                window.location.href = '/submitted';
            }
        })
        .catch(() => {});
}, 5000);

function updateTimerDisplay() {
    const mins = Math.floor(examDuration / 60);
    const secs = examDuration % 60;
    const timerSpan = document.querySelector('#timer .timer-time');
    if (timerSpan) {
        timerSpan.textContent = `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
    }
}

function renderPalette() {
    desktopGrid = document.getElementById('paletteGrid');
    mobileGrid = document.getElementById('mobilePaletteGrid');
    
    if (desktopGrid) desktopGrid.innerHTML = '';
    if (mobileGrid) mobileGrid.innerHTML = '';
    
    questions.forEach((q, idx) => {
        if (desktopGrid) {
            const btn = document.createElement('div');
            btn.className = 'palette-btn';
            btn.textContent = idx + 1;
            btn.addEventListener('click', () => {
                loadQuestion(idx);
                if (paletteOverlay) paletteOverlay.classList.remove('active');
                if (mobilePalette) mobilePalette.classList.remove('active');
            });
            desktopGrid.appendChild(btn);
            q.btnRef = btn;
        }
        
        if (mobileGrid) {
            const mBtn = document.createElement('div');
            mBtn.className = 'palette-btn';
            mBtn.textContent = idx + 1;
            mBtn.addEventListener('click', () => {
                loadQuestion(idx);
                if (paletteOverlay) paletteOverlay.classList.remove('active');
                if (mobilePalette) mobilePalette.classList.remove('active');
            });
            mobileGrid.appendChild(mBtn);
            q.mobileBtnRef = mBtn;
        }
        
        updatePaletteButton(idx);
    });
}

function updatePaletteButton(index) {
    const q = questions[index];
    if (!q) return;
    
    const btn = q.btnRef;
    const mBtn = q.mobileBtnRef;
    
    function updateBtn(b) {
        if (!b) return;
        b.classList.remove('not-visited', 'not-answered', 'answered', 'marked');
        if (!visited.has(index)) {
            b.classList.add('not-visited');
        } else if (markedForReview.has(index)) {
            b.classList.add('marked');
        } else if (responses[index] !== undefined && responses[index] !== null && responses[index] !== '') {
            b.classList.add('answered');
        } else {
            b.classList.add('not-answered');
        }
    }
    
    updateBtn(btn);
    updateBtn(mBtn);
}

function refreshQuestionStatus(index) {
    const status = markedForReview.has(index)
        ? 'Marked for Review'
        : (responses[index] ? 'Answered' : 'Not Answered');
    const statusEl = document.getElementById('currentStatus');
    if (statusEl) statusEl.textContent = status;
}

function refreshQuestionCounters(index) {
    const headerNumber = document.getElementById('questionNumberHeader');
    const panelNumber = document.getElementById('questionNumberPanel');
    if (headerNumber) headerNumber.textContent = (index + 1).toString();
    if (panelNumber) panelNumber.textContent = (index + 1).toString();
}

function loadQuestion(index) {
    currentIndex = index;
    visited.add(index);
    const q = questions[index];

    refreshQuestionCounters(index);
    refreshQuestionStatus(index);

    // ── Question text ─────────────────────────────────────────────────────────
    const qText = document.getElementById('questionText');
    if (qText) {
        qText.id = 'questionText'; // keep id stable
        if (window.RRBMultimodal) {
            window.RRBMultimodal.render(qText, q).catch(() => {
                qText.textContent = 'Q' + (index + 1) + ': ' + (q.question || '');
            });
        } else {
            qText.innerHTML = q.question || '';
        }
    }

    // ── Static image (database image_path) ───────────────────────────────────
    let imageContainer = document.getElementById('questionImageContainer');
    const questionPanel = document.querySelector('.question-panel');
    if (!imageContainer && questionPanel) {
        imageContainer = document.createElement('div');
        imageContainer.id = 'questionImageContainer';
        imageContainer.style.cssText = 'margin-bottom:16px;text-align:center;';
        if (qText) qText.parentNode.insertBefore(imageContainer, qText.nextSibling);
    }
    if (imageContainer) {
        if (q.image_path) {
            imageContainer.innerHTML =
                `<img src="/static/${q.image_path}" alt="Question diagram"
                      style="max-width:100%;max-height:280px;border-radius:8px;
                             box-shadow:0 4px 12px rgba(0,0,0,.12);cursor:zoom-in;"
                      onclick="if(window.RRBMultimodal)RRBMultimodal.openLightbox(this.src,'Question diagram')">`;
        } else {
            imageContainer.innerHTML = '';
        }
    }

    // ── Options ───────────────────────────────────────────────────────────────
    const optsDiv = document.getElementById('optionsContainer');
    if (optsDiv) {
        optsDiv.innerHTML = '';
        let anyOptMath = false;
        ['A', 'B', 'C', 'D'].forEach(opt => {
            const key     = `option_${opt.toLowerCase()}`;
            const optText = (q[key] || '').trim();
            if (!optText) return;

            const hasMath = window.RRBMultimodal
                ? window.RRBMultimodal.needsMath(optText)
                : optText.includes('$') || optText.includes('\\ce{');

            if (hasMath) anyOptMath = true;

            const div = document.createElement('div');
            div.className = 'option-item';
            if (responses[index] === opt) div.classList.add('selected');

            const label = document.createElement('span');
            label.style.cssText = 'flex:1;line-height:1.7;';
            label.className = hasMath ? 'mj-content' : '';
            label.innerHTML  = `<strong>${opt}.</strong>&nbsp;${optText}`;

            const radio = document.createElement('input');
            radio.type  = 'radio';
            radio.name  = 'option';
            radio.value = opt;
            if (responses[index] === opt) radio.checked = true;

            div.appendChild(radio);
            div.appendChild(label);
            div.addEventListener('click', function () {
                this.querySelector('input[type=radio]').checked = true;
                selectOption(index, opt);
            });
            optsDiv.appendChild(div);
        });

        // Typeset options that contain math
        if (anyOptMath) {
            if (window.RRBMultimodal) {
                window.RRBMultimodal.initMathJax().then(() => {
                    window.RRBMultimodal.typeset(optsDiv);
                });
            } else if (window.MathJax?.typesetPromise) {
                MathJax.typesetPromise([optsDiv]).catch(() => {});
            }
        }
    }

    questions.forEach((_, i) => updatePaletteButton(i));
}

function selectOption(qIndex, option) {
    responses[qIndex] = option;
    markedForReview.delete(qIndex);
    refreshQuestionStatus(qIndex);
    
    fetch('/save_answer', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            question_id: questions[qIndex].id,
            selected_option: option
        })
    });
    
    const optionsDiv = document.getElementById('optionsContainer');
    if (optionsDiv) {
        Array.from(optionsDiv.children).forEach((child) => {
            const radio = child.querySelector('input[type="radio"]');
            if (radio && radio.value === option) {
                child.classList.add('selected');
                radio.checked = true;
            } else {
                child.classList.remove('selected');
            }
        });
    }
    
    updatePaletteButton(qIndex);
}

function saveCurrentAndNext() {
    const selectedRadio = document.querySelector('input[name="option"]:checked');
    if (selectedRadio) {
        const option = selectedRadio.value;
        if (responses[currentIndex] !== option) {
            selectOption(currentIndex, option);
        }
    }
    
    if (currentIndex < questions.length - 1) {
        loadQuestion(currentIndex + 1);
    } else {
        alert('This is the last question.');
    }
}

function markForReviewAndNext() {
    markedForReview.add(currentIndex);
    updatePaletteButton(currentIndex);
    saveCurrentAndNext();
}

function clearResponse() {
    delete responses[currentIndex];
    markedForReview.delete(currentIndex);
    
    fetch('/save_answer', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
            question_id: questions[currentIndex].id,
            selected_option: ''
        })
    });
    
    const optionsDiv = document.getElementById('optionsContainer');
    if (optionsDiv) {
        Array.from(optionsDiv.children).forEach(child => {
            child.classList.remove('selected');
            const radio = child.querySelector('input[type="radio"]');
            if (radio) radio.checked = false;
        });
    }
    
    refreshQuestionStatus(currentIndex);
    updatePaletteButton(currentIndex);
}

function silentSubmit() {
    sessionStorage.setItem('examSubmitted', 'true');
    fetch('/submit_exam', {method: 'POST'})
        .then(() => {
            if (document.exitFullscreen){
                window.location.href = '/submitted';
            }
        });
}

function submitExam() {
    if (confirm('Are you sure you want to submit?')) {
        silentSubmit();
    }
}

// Event listeners for buttons
document.getElementById('saveNextBtn')?.addEventListener('click', saveCurrentAndNext);
document.getElementById('markReviewBtn')?.addEventListener('click', markForReviewAndNext);
document.getElementById('clearBtn')?.addEventListener('click', clearResponse);
document.getElementById('submitBtn')?.addEventListener('click', submitExam);
// ================================================================
// FEATURE #1 — Hindi/English Translation Toggle
// ================================================================

let translationState = 'en';          // 'en' or 'hi'
let translationCache = {};             // { questionIndex: { en: {...}, hi: {...} } }
let translationInProgress = false;

function injectTranslateButton() {
    if (document.getElementById('translateBtn')) return;
    const btn = document.createElement('button');
    btn.id = 'translateBtn';
    btn.className = 'btn-translate';
    btn.title = 'Translate question';
    btn.innerHTML = '<span id="translateLabel">हिंदी</span>';
    btn.onclick = toggleTranslation;

    // Insert into exam header area
    const header = document.querySelector('.exam-header');
    if (header) header.appendChild(btn);

    // Inject button styles
    const st = document.createElement('style');
    st.textContent = `
        .btn-translate {
            background: rgba(255,255,255,0.18);
            color: white;
            border: 1.5px solid rgba(255,255,255,0.5);
            padding: 5px 14px;
            border-radius: 16px;
            font-family: 'Poppins', 'Noto Sans Devanagari', sans-serif;
            font-size: 0.85rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.25s;
            white-space: nowrap;
            display: flex;
            align-items: center;
            gap: 5px;
        }
        .btn-translate::before { content: '🌐'; font-size: 0.9rem; }
        .btn-translate:hover { background: rgba(255,255,255,0.32); }
        .btn-translate.loading { opacity: 0.6; cursor: wait; }
        .translation-badge {
            display: inline-block;
            background: rgba(255,255,255,0.2);
            color: white;
            font-size: 0.72rem;
            padding: 1px 7px;
            border-radius: 10px;
            margin-left: 6px;
            vertical-align: middle;
        }
    `;
    document.head.appendChild(st);
}

async function toggleTranslation() {
    if (translationInProgress) return;
    const q = questions[currentIndex];
    if (!q) return;

    // If currently in Hindi → switch back to English (from cache, instant)
    if (translationState === 'hi') {
        translationState = 'en';
        document.getElementById('translateLabel').textContent = 'हिंदी';
        restoreEnglish(currentIndex);
        return;
    }

    // If Hindi already cached for this question → use it
    if (translationCache[currentIndex]?.hi) {
        translationState = 'hi';
        document.getElementById('translateLabel').textContent = 'ENGLISH';
        applyTranslation(currentIndex, translationCache[currentIndex].hi);
        return;
    }

    // Fetch translation via Gemini
    translationInProgress = true;
    const btn = document.getElementById('translateBtn');
    btn.classList.add('loading');
    document.getElementById('translateLabel').textContent = '...';

    try {
        const apiKey = window._GEMINI_KEY || '';
        if (!apiKey) throw new Error('API key not configured');

        const textsToTranslate = {
            question: q.question || '',
            option_a: q.option_a || '',
            option_b: q.option_b || '',
            option_c: q.option_c || '',
            option_d: q.option_d || '',
        };

        const prompt = `Translate the following exam question and options from English to Hindi.
Return ONLY a valid JSON object with these exact keys: question, option_a, option_b, option_c, option_d
Do not translate mathematical expressions, formulas, or scientific terms — keep them as-is.
Do not add any explanation or markdown.

Input JSON:
${JSON.stringify(textsToTranslate)}`;

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.1, maxOutputTokens: 1024 }
            })
        });

        const data = await res.json();
        let raw = data?.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '';

        // Strip markdown fences
        raw = raw.replace(/^```(?:json)?|```$/gm, '').trim();
        const translated = JSON.parse(raw);

        // Cache it
        translationCache[currentIndex] = { hi: translated };
        translationState = 'hi';
        document.getElementById('translateLabel').textContent = 'ENGLISH';
        applyTranslation(currentIndex, translated);

    } catch (err) {
        console.warn('Translation error:', err);
        document.getElementById('translateLabel').textContent = 'हिंदी';
        // Show user-friendly error
        const qText = document.getElementById('questionText');
        if (qText) {
            const errMsg = document.createElement('div');
            errMsg.style.cssText = 'background:#fde;color:#c00;padding:6px 12px;border-radius:6px;font-size:0.82rem;margin-top:8px;';
            errMsg.textContent = '⚠ Translation failed. Check API key in apikey.env.';
            qText.appendChild(errMsg);
            setTimeout(() => errMsg.remove(), 4000);
        }
    } finally {
        translationInProgress = false;
        btn.classList.remove('loading');
    }
}

function applyTranslation(index, translated) {
    const q = questions[index];
    const qText = document.getElementById('questionText');
    const optsDiv = document.getElementById('optionsContainer');

    // Update question text
    if (qText) {
        qText.innerHTML = `<span style="font-family:'Noto Sans Devanagari',Arial,sans-serif;">`
            + `Q${index + 1}: ${translated.question}</span>`
            + `<span class="translation-badge">हिंदी</span>`;
    }

    // Update options
    if (optsDiv) {
        ['A', 'B', 'C', 'D'].forEach(opt => {
            const key  = `option_${opt.toLowerCase()}`;
            const text = translated[key];
            if (!text) return;
            const el   = optsDiv.querySelectorAll('.option-item')[['A','B','C','D'].indexOf(opt)];
            if (el) {
                const span = el.querySelector('span');
                if (span) span.style.fontFamily = "'Noto Sans Devanagari', Arial, sans-serif";
                if (span) span.innerHTML = `${opt}. ${text}`;
            }
        });
    }
}

function restoreEnglish(index) {
    const q = questions[index];
    const qText = document.getElementById('questionText');
    if (qText) {
        qText.style.fontFamily = '';
        qText.innerHTML = `Q${index + 1}: ${q.question}`;
        // Re-render math if needed
        if (window.RRBMultimodal && ((q.question || '').includes('$') || (q.question || '').includes('\\('))) {
            window.RRBMultimodal.renderMath(qText, `Q${index + 1}: ${q.question}`);
        }
    }
    // Re-render full question options via loadQuestion
    const optsDiv = document.getElementById('optionsContainer');
    if (optsDiv) {
        ['A', 'B', 'C', 'D'].forEach((opt, idx) => {
            const text = q[`option_${opt.toLowerCase()}`];
            const el   = optsDiv.querySelectorAll('.option-item')[idx];
            if (el && text) {
                const span = el.querySelector('span');
                if (span) { span.style.fontFamily = ''; span.innerHTML = `${opt}. ${text}`; }
            }
        });
    }
}

// Reset translation when navigating to another question
const _origLoadQuestion = loadQuestion;
window.loadQuestion = function(index) {
    // Reset translation state when changing questions
    if (index !== currentIndex) {
        translationState = 'en';
        const lbl = document.getElementById('translateLabel');
        if (lbl) lbl.textContent = 'हिंदी';
    }
    _origLoadQuestion(index);
};

// Read Gemini key from meta tag (injected by server)
document.addEventListener('DOMContentLoaded', () => {
    const meta = document.querySelector('meta[name="gemini-key"]');
    if (meta) window._GEMINI_KEY = meta.content;
    injectTranslateButton();
});
// Also run immediately if DOM already loaded
if (document.readyState !== 'loading') {
    const meta = document.querySelector('meta[name="gemini-key"]');
    if (meta) window._GEMINI_KEY = meta.content;
    setTimeout(injectTranslateButton, 200);
}
