/**
 * Secure Exam Taking JavaScript Controller
 * Handles server timer, periodic auto-save, question navigation, and flag state.
 */

class ExamController {
    constructor(attemptId, totalSeconds, questionsCount) {
        this.attemptId = attemptId;
        this.remainingSeconds = totalSeconds;
        this.questionsCount = questionsCount;
        this.currentQuestionIndex = 0;
        this.flaggedQuestions = new Set();
        this.timerInterval = null;
        this.autoSaveInterval = null;

        this.init();
    }

    init() {
        this.startTimer();
        this.showQuestion(0);
        this.setupAutoSave();
        this.setupIntegrityMonitor();
        this.bindEvents();
    }

    startTimer() {
        const timerDisplay = document.getElementById('timer-display');
        
        const updateDisplay = () => {
            if (this.remainingSeconds <= 0) {
                clearInterval(this.timerInterval);
                if (timerDisplay) timerDisplay.textContent = "00:00:00";
                alert("Time is up! Submitting your exam now.");
                document.getElementById('exam-form').submit();
                return;
            }

            const hrs = Math.floor(this.remainingSeconds / 3600);
            const mins = Math.floor((this.remainingSeconds % 3600) / 60);
            const secs = this.remainingSeconds % 60;

            const formatted = [
                hrs.toString().padStart(2, '0'),
                mins.toString().padStart(2, '0'),
                secs.toString().padStart(2, '0')
            ].join(':');

            if (timerDisplay) {
                timerDisplay.textContent = formatted;
                if (this.remainingSeconds < 300) { // Under 5 mins warning color
                    timerDisplay.style.color = '#ef4444';
                }
            }

            this.remainingSeconds--;
        };

        updateDisplay();
        this.timerInterval = setInterval(updateDisplay, 1000);

        // Periodically sync timer with server every 30 seconds
        setInterval(() => {
            fetch(`/api/exam/timer/${this.attemptId}`)
                .then(r => r.json())
                .then(data => {
                    if (data.remaining_seconds !== undefined) {
                        this.remainingSeconds = data.remaining_seconds;
                    }
                })
                .catch(err => console.warn("Timer sync failed:", err));
        }, 30000);
    }

    showQuestion(index) {
        if (index < 0 || index >= this.questionsCount) return;

        // Hide all question blocks
        document.querySelectorAll('.question-block').forEach((el, i) => {
            el.style.display = (i === index) ? 'block' : 'none';
        });

        this.currentQuestionIndex = index;

        // Update nav button active classes
        document.querySelectorAll('.q-nav-btn').forEach((btn, i) => {
            if (i === index) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        // Update Prev/Next buttons
        const prevBtn = document.getElementById('btn-prev');
        const nextBtn = document.getElementById('btn-next');
        if (prevBtn) prevBtn.disabled = (index === 0);
        if (nextBtn) nextBtn.disabled = (index === this.questionsCount - 1);
    }

    toggleFlag(questionId, index) {
        const btn = document.querySelector(`.q-nav-btn[data-index="${index}"]`);
        if (this.flaggedQuestions.has(questionId)) {
            this.flaggedQuestions.delete(questionId);
            if (btn) btn.classList.remove('flagged');
        } else {
            this.flaggedQuestions.add(questionId);
            if (btn) btn.classList.add('flagged');
        }
    }

    saveCurrentAnswer(questionId) {
        const container = document.querySelector(`.question-block[data-question-id="${questionId}"]`);
        if (!container) return;

        let responseVal = null;
        const inputs = container.querySelectorAll('input, textarea');

        if (inputs.length > 0) {
            const firstType = inputs[0].type;

            if (firstType === 'radio') {
                const checked = container.querySelector('input[type="radio"]:checked');
                if (checked) responseVal = checked.value;
            } else if (firstType === 'checkbox') {
                const checkedList = Array.from(container.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
                responseVal = checkedList;
            } else if (firstType === 'textarea' || firstType === 'text') {
                responseVal = inputs[0].value.trim();
            }
        }

        // Highlight nav button as answered
        const navBtn = document.querySelector(`.q-nav-btn[data-q-id="${questionId}"]`);
        if (navBtn) {
            if (responseVal && (Array.isArray(responseVal) ? responseVal.length > 0 : true)) {
                navBtn.classList.add('answered');
            } else {
                navBtn.classList.remove('answered');
            }
        }

        // Background POST to API
        fetch('/api/exam/save-answer', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                attempt_id: this.attemptId,
                question_id: questionId,
                response: responseVal
            })
        })
        .then(r => r.json())
        .then(data => {
            const statusIndicator = document.getElementById('save-indicator');
            if (statusIndicator) {
                statusIndicator.textContent = "Saved";
                setTimeout(() => { statusIndicator.textContent = ""; }, 2000);
            }
        })
        .catch(err => console.error("Auto-save failed:", err));
    }

    setupAutoSave() {
        // Auto save on any input change
        document.querySelectorAll('.question-block input, .question-block textarea').forEach(input => {
            input.addEventListener('change', (e) => {
                const qBlock = e.target.closest('.question-block');
                if (qBlock) {
                    const qId = parseInt(qBlock.dataset.questionId);
                    this.saveCurrentAnswer(qId);
                }
            });
        });
    }

    setupIntegrityMonitor() {
        let lastBlurTime = 0;

        const reportViolation = (eventType, details) => {
            const now = Date.now();
            if (now - lastBlurTime < 3000) return;
            lastBlurTime = now;

            this.showSecurityToast(`⚠️ Proctoring Alert: ${details}`);

            fetch('/api/exam/integrity-event', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    attempt_id: this.attemptId,
                    event_type: eventType,
                    details: details
                })
            }).then(r => r.json()).then(data => {
                if (data && data.integrity_score !== undefined) {
                    const badge = document.getElementById('integrity-badge');
                    if (badge) {
                        badge.textContent = `Integrity: ${data.integrity_score}%`;
                        if (data.integrity_score < 75) badge.style.backgroundColor = 'rgba(239, 68, 68, 0.25)';
                    }
                }
            }).catch(() => {});
        };

        window.addEventListener('blur', () => {
            reportViolation('window_blur', 'Window lost focus or application switched.');
        });

        document.addEventListener('visibilitychange', () => {
            if (document.hidden) {
                reportViolation('tab_switch', 'Browser tab changed or minimized.');
            }
        });

        document.addEventListener('copy', () => {
            reportViolation('copy_attempt', 'Attempted to copy exam question content.');
        });
    }

    showSecurityToast(message) {
        let toast = document.getElementById('proctor-toast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'proctor-toast';
            toast.style.cssText = 'position:fixed;bottom:24px;right:24px;background:rgba(239,68,68,0.95);color:#fff;padding:12px 20px;border-radius:10px;font-size:0.85rem;font-weight:600;box-shadow:0 10px 25px rgba(0,0,0,0.4);z-index:9999;transition:all 0.3s ease;display:flex;align-items:center;gap:10px;backdrop-filter:blur(8px);';
            document.body.appendChild(toast);
        }
        toast.textContent = message;
        toast.style.opacity = '1';
        toast.style.transform = 'translateY(0)';
        clearTimeout(this._toastTimeout);
        this._toastTimeout = setTimeout(() => {
            if (toast) {
                toast.style.opacity = '0';
                toast.style.transform = 'translateY(20px)';
            }
        }, 4000);
    }

    bindEvents() {
        const prevBtn = document.getElementById('btn-prev');
        const nextBtn = document.getElementById('btn-next');

        if (prevBtn) {
            prevBtn.addEventListener('click', () => {
                this.showQuestion(this.currentQuestionIndex - 1);
            });
        }

        if (nextBtn) {
            nextBtn.addEventListener('click', () => {
                this.showQuestion(this.currentQuestionIndex + 1);
            });
        }

        document.querySelectorAll('.btn-flag').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const qId = parseInt(e.target.dataset.questionId);
                const idx = parseInt(e.target.dataset.index);
                this.toggleFlag(qId, idx);
            });
        });

        document.querySelectorAll('.q-nav-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.target.dataset.index);
                this.showQuestion(idx);
            });
        });
    }
}
