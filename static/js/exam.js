// ================================================================
// IN-SCREEN TOAST NOTIFICATIONS (No Fullscreen Exit / No Window Blur)
// ================================================================
function showToastNotification(message, type = 'info') {
    let container = document.getElementById('examToastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'examToastContainer';
        container.style.cssText = `
            position: fixed;
            top: 60px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 999999;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
            pointer-events: none;
        `;
        document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.style.cssText = `
        background: ${type === 'warning' ? '#be123c' : 'linear-gradient(135deg, #0f172a, #1e293b)'};
        color: #ffffff;
        padding: 10px 22px;
        border-radius: 30px;
        font-size: 0.88rem;
        font-weight: 600;
        box-shadow: 0 10px 30px rgba(0,0,0,0.35);
        display: flex;
        align-items: center;
        gap: 8px;
        border: 1px solid rgba(255,255,255,0.2);
        pointer-events: auto;
    `;
    toast.innerHTML = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.4s ease';
        setTimeout(() => toast.remove(), 400);
    }, 3500);
}

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


// Detect mobile device (strictly mobile UAs, excluding touchscreen Windows PCs/laptops)
function isMobileDevice() {
    const ua = navigator.userAgent || '';
    return /Android|webOS|iPhone|iPod|BlackBerry|IEMobile|Opera Mini/i.test(ua);
}

const isMobile = isMobileDevice();
let fullscreenEnabled = false;

// ── GLOBAL TAB SWITCH & FULLSCREEN SECURITY MONITOR ──────────────────────────
let _submittingTriggered = false;

function triggerAutoSubmit(reason) {
    if (_submittingTriggered) return;
    if (sessionStorage.getItem('examSubmitted')) return;
    if (window._isSubmittingModalOpen) return;

    _submittingTriggered = true;
    console.warn('Security violation auto-submit:', reason);
    silentSubmit(reason, true);
}

// Global Tab Switch & Blur listeners (enforced on all desktop browsers)
document.addEventListener('visibilitychange', () => {
    if (document.hidden || document.visibilityState === 'hidden') {
        triggerAutoSubmit('tab_switch');
    }
});

window.addEventListener('blur', () => {
    triggerAutoSubmit('tab_switch');
});

// Fullscreen exit monitor for desktop devices
if (!isMobile) {
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
                showFullscreenPrompt(true);
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
                <h3 style="color:#b71c1c;margin-bottom:10px;">Fullscreen Required</h3>
                <p style="color:#555;margin-bottom:20px;line-height:1.6;">
                    This exam must be taken in fullscreen mode to continue.
                </p>
                <button id="fsBtnRe" style="background:#1b5e20;color:white;border:none;
                        padding:12px 28px;border-radius:30px;font-size:1rem;cursor:pointer;
                        display:inline-flex;align-items:center;gap:8px;">
                    ⛶ Enter Fullscreen
                </button>
            </div>`;
        document.body.appendChild(div);
        document.getElementById('fsBtnRe').addEventListener('click', () => {
            requestFullscreen();
            setTimeout(() => {
                if (isInFullscreen()) {
                    fullscreenEnabled = true;
                    const p = document.getElementById('fullscreenPrompt');
                    if (p) p.remove();
                }
            }, 400);
        });
    }

    function exitHandler() {
        if (isInFullscreen()) {
            const p = document.getElementById('fullscreenPrompt');
            if (p) p.remove();
            return;
        }

        if (!isInFullscreen() && !sessionStorage.getItem('examSubmitted') && !window._isSubmittingModalOpen) {
            triggerAutoSubmit('fullscreen_exit');
        }
    }

    enterFullscreen();

    document.addEventListener('fullscreenchange',       exitHandler);
    document.addEventListener('webkitfullscreenchange', exitHandler);
    document.addEventListener('mozfullscreenchange',    exitHandler);
    document.addEventListener('MSFullscreenChange',     exitHandler);
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
    const hasAns = responses[index] !== undefined && responses[index] !== null && responses[index] !== '';
    const isMarked = markedForReview.has(index);
    let status = 'Not Answered';
    if (isMarked && hasAns) {
        status = 'Answered & Marked for Review';
    } else if (isMarked) {
        status = 'Marked for Review';
    } else if (hasAns) {
        status = 'Answered';
    }
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

    // In-screen notification on reaching the last question
    if (questions.length > 0 && index === questions.length - 1 && !window._lastQNotified) {
        window._lastQNotified = true;
        showToastNotification('📌 Note: You are on the last question of the exam.', 'info');
    } else if (index < questions.length - 1) {
        window._lastQNotified = false;
    }

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

    if (translationState === 'hi') {
        if (q.question_hi && q.question_hi.trim() !== '') {
            applyTranslation(index, {
                question: q.question_hi,
                option_a: q.option_a_hi || q.option_a,
                option_b: q.option_b_hi || q.option_b,
                option_c: q.option_c_hi || q.option_c,
                option_d: q.option_d_hi || q.option_d
            });
        } else if (translationCache[index]?.hi) {
            applyTranslation(index, translationCache[index].hi);
        }
    }

    questions.forEach((_, i) => updatePaletteButton(i));
}

function selectOption(qIndex, option, keepMarked = false) {
    responses[qIndex] = option;
    if (!keepMarked) {
        markedForReview.delete(qIndex);
    }
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
        selectOption(currentIndex, option, false);
    }
    
    if (currentIndex < questions.length - 1) {
        loadQuestion(currentIndex + 1);
    } else {
        showToastNotification('📌 Note: You are on the last question of the exam.', 'info');
    }
}

function markForReviewAndNext() {
    markedForReview.add(currentIndex);
    const selectedRadio = document.querySelector('input[name="option"]:checked');
    if (selectedRadio) {
        const option = selectedRadio.value;
        selectOption(currentIndex, option, true);
    } else {
        refreshQuestionStatus(currentIndex);
        updatePaletteButton(currentIndex);
    }
    
    if (currentIndex < questions.length - 1) {
        loadQuestion(currentIndex + 1);
    } else {
        showToastNotification('📌 Note: You are on the last question of the exam.', 'info');
    }
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

function silentSubmit(reason, force = false) {
    if (sessionStorage.getItem('examSubmitted') && !force) return;
    sessionStorage.setItem('examSubmitted', 'true');

    let subOverlay = document.getElementById('submittingOverlay');
    if (!subOverlay) {
        subOverlay = document.createElement('div');
        subOverlay.id = 'submittingOverlay';
        subOverlay.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background: rgba(15, 23, 42, 0.92); z-index: 9999999;
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            color: #ffffff; font-family: 'Poppins', sans-serif;
        `;
        subOverlay.innerHTML = `
            <div style="font-size: 2.5rem; margin-bottom: 12px;">⏳</div>
            <h3 style="margin-bottom: 8px;">Submitting Exam...</h3>
            <p style="font-size: 0.85rem; opacity: 0.8;">Please wait while your answers are recorded.</p>
        `;
        document.body.appendChild(subOverlay);
    }

    fetch('/submit_exam', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reason: reason || 'manual' })
    })
    .then(res => {
        if (!res.ok) throw new Error('Submission response error');
        return res.json();
    })
    .then(data => {
        window.location.href = data.redirect || '/submitted';
    })
    .catch(err => {
        console.warn('Submit error fallback:', err);
        window.location.href = '/submitted';
    });
}

function showSubmitConfirmationModal() {
    if (document.getElementById('submitConfirmModal')) return;
    
    window._isSubmittingModalOpen = true;

    // Count statistics
    let answered = 0;
    let marked = markedForReview.size;
    let notVisited = 0;
    let notAnswered = 0;

    questions.forEach((q, idx) => {
        if (responses[idx]) {
            answered++;
        } else if (visited.has(idx)) {
            notAnswered++;
        } else {
            notVisited++;
        }
    });

    const total = questions.length;

    const modal = document.createElement('div');
    modal.id = 'submitConfirmModal';
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background: rgba(15, 23, 42, 0.82);
        backdrop-filter: blur(8px);
        z-index: 999999;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
    `;

    modal.innerHTML = `
        <div style="background: #ffffff; width: 100%; max-width: 440px; border-radius: 20px;
                    padding: 28px 24px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.4);
                    text-align: center; font-family: 'Poppins', sans-serif;">
            <div style="width: 60px; height: 60px; background: #e0e7ff; color: #3730a3;
                        border-radius: 50%; display: flex; align-items: center; justify-content: center;
                        margin: 0 auto 16px; font-size: 1.8rem;">
                📝
            </div>
            <h3 style="color: #1e1b4b; font-size: 1.3rem; font-weight: 700; margin-bottom: 8px;">
                Are you sure you want to submit your exam?
            </h3>
            <p style="color: #64748b; font-size: 0.88rem; margin-bottom: 20px;">
                Please review your exam summary before final submission.
            </p>

            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 24px; text-align: left;">
                <div style="background: #f8fafc; padding: 10px 14px; border-radius: 12px; border: 1px solid #e2e8f0;">
                    <div style="font-size: 0.72rem; color: #64748b; font-weight: 600;">ANSWERED</div>
                    <div style="font-size: 1.1rem; color: #16a34a; font-weight: 700;">${answered} / ${total}</div>
                </div>
                <div style="background: #f8fafc; padding: 10px 14px; border-radius: 12px; border: 1px solid #e2e8f0;">
                    <div style="font-size: 0.72rem; color: #64748b; font-weight: 600;">MARKED FOR REVIEW</div>
                    <div style="font-size: 1.1rem; color: #7c3aed; font-weight: 700;">${marked}</div>
                </div>
                <div style="background: #f8fafc; padding: 10px 14px; border-radius: 12px; border: 1px solid #e2e8f0;">
                    <div style="font-size: 0.72rem; color: #64748b; font-weight: 600;">NOT ANSWERED</div>
                    <div style="font-size: 1.1rem; color: #dc2626; font-weight: 700;">${notAnswered}</div>
                </div>
                <div style="background: #f8fafc; padding: 10px 14px; border-radius: 12px; border: 1px solid #e2e8f0;">
                    <div style="font-size: 0.72rem; color: #64748b; font-weight: 600;">UNVISITED</div>
                    <div style="font-size: 1.1rem; color: #64748b; font-weight: 700;">${notVisited}</div>
                </div>
            </div>

            <div style="display: flex; gap: 12px;">
                <button id="cancelSubmitBtn" style="flex: 1; padding: 12px; border: 1.5px solid #cbd5e1;
                        background: #ffffff; color: #475569; border-radius: 12px; font-weight: 600;
                        font-size: 0.92rem; cursor: pointer; transition: all 0.2s;">
                    Cancel
                </button>
                <button id="confirmSubmitBtn" style="flex: 1; padding: 12px; border: none;
                        background: linear-gradient(135deg, #16a34a, #15803d); color: #ffffff;
                        border-radius: 12px; font-weight: 600; font-size: 0.92rem; cursor: pointer;
                        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.3); transition: all 0.2s;">
                    Yes, Submit
                </button>
            </div>
        </div>
    `;

    document.body.appendChild(modal);

    document.getElementById('cancelSubmitBtn').addEventListener('click', () => {
        modal.remove();
        window._isSubmittingModalOpen = false;
    });

    document.getElementById('confirmSubmitBtn').addEventListener('click', () => {
        modal.remove();
        window._isSubmittingModalOpen = false;
        silentSubmit('manual', true);
    });
}

function submitExam() {
    showSubmitConfirmationModal();
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

    // 1. Check if database already has Hindi text for this question (q.question_hi)
    if (q.question_hi && q.question_hi.trim() !== '') {
        const hiObj = {
            question: q.question_hi,
            option_a: q.option_a_hi || q.option_a,
            option_b: q.option_b_hi || q.option_b,
            option_c: q.option_c_hi || q.option_c,
            option_d: q.option_d_hi || q.option_d
        };
        translationCache[currentIndex] = translationCache[currentIndex] || {};
        translationCache[currentIndex].hi = hiObj;
        translationState = 'hi';
        document.getElementById('translateLabel').textContent = 'ENGLISH';
        applyTranslation(currentIndex, hiObj);
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

        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
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
