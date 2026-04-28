document.addEventListener('DOMContentLoaded', () => {
    // --- Mode Toggling ---
    const minimizeBtn = document.getElementById('minimizeDashboardBtn');
    const maximizeBtn = document.getElementById('maximizeWidgetBtn');
    
    if (minimizeBtn) {
        minimizeBtn.addEventListener('click', () => {
            document.body.classList.remove('dashboard-mode');
            document.body.classList.add('widget-mode');
        });
    }
    
    if (maximizeBtn) {
        maximizeBtn.addEventListener('click', () => {
            document.body.classList.remove('widget-mode');
            document.body.classList.add('dashboard-mode');
            const widgetView = document.getElementById('widget-view');
            if (widgetView) widgetView.classList.remove('minimized');
        });
    }

    // --- Widget Minimize/Maximize Logic ---
    const minimizeWidgetBtn = document.getElementById('minimizeWidgetBtn');
    const miniAssistantBtn = document.getElementById('miniAssistant');
    const widgetView = document.getElementById('widget-view');

    if (minimizeWidgetBtn) {
        minimizeWidgetBtn.addEventListener('click', () => {
            if (widgetView) widgetView.classList.add('minimized');
        });
    }

    let hasDraggedWidget = false;

    if (miniAssistantBtn) {
        miniAssistantBtn.addEventListener('click', (e) => {
            if (hasDraggedWidget) {
                // If we dragged, don't trigger the click action
                e.preventDefault();
                e.stopPropagation();
                hasDraggedWidget = false;
                return;
            }
            if (widgetView) widgetView.classList.remove('minimized');
        });
    }

    // --- DOM Elements (Dashboard) ---
    const micRing = document.getElementById('mic-ring');
    const modeState = document.getElementById('mode-state');
    const currentTask = document.getElementById('current-task');
    const wakeWordEl = document.getElementById('wake-word');
    const stopPhraseEl = document.getElementById('stop-phrase');
    const lastTranscription = document.getElementById('last-transcription');
    const lastAction = document.getElementById('last-action');
    const logContainer = document.getElementById('log-container');

    // --- DOM Elements (Widget) ---
    const widgetStatusDot = document.getElementById('widget-statusDot');
    const widgetStatusText = document.getElementById('widget-statusText');
    const widgetTranscriptText = document.getElementById('widget-transcriptText');
    const widgetLogFeed = document.getElementById('widget-logFeed');
    let agentState = 'sleeping';

    function setWidgetState(state) {
        agentState = state;
        const stateConfig = {
            sleeping: { dot: 'sleeping', text: 'Awaiting Wake Word' },
            listening: { dot: 'listening', text: 'Listening…' },
            processing: { dot: 'processing', text: 'Processing…' },
            speaking: { dot: 'speaking', text: 'Executing…' }
        };
        const cfg = stateConfig[state] || stateConfig.sleeping;
        if (widgetStatusDot) widgetStatusDot.className = `status-indicator ${cfg.dot}`;
        if (widgetStatusText) widgetStatusText.textContent = cfg.text;
    }

    function addWidgetLog(text) {
        if (!widgetLogFeed) return;
        const item = document.createElement('div');
        item.className = 'log-item new';
        item.textContent = `› ${text}`;
        widgetLogFeed.prepend(item);
        setTimeout(() => item.classList.remove('new'), 200);
        while (widgetLogFeed.children.length > 3) widgetLogFeed.removeChild(widgetLogFeed.lastChild);
    }

    // --- Setup SSE Connection ---
    const evtSource = new EventSource("/events");

    function appendLog(type, message) {
        if (!logContainer) return;
        const div = document.createElement('div');
        div.className = `log-entry ${type}`;
        const now = new Date();
        const timeStr = now.toLocaleTimeString([], { hour12: false });
        div.innerHTML = `<span class="timestamp">[${timeStr}]</span> ${message}`;
        logContainer.appendChild(div);
        logContainer.scrollTop = logContainer.scrollHeight;
    }

    evtSource.onmessage = function(event) {
        try {
            const data = JSON.parse(event.data);
            const type = data.type;
            const payload = data.payload;

            switch(type) {
                case 'init':
                    if (wakeWordEl) wakeWordEl.textContent = payload.wake_word;
                    if (stopPhraseEl) stopPhraseEl.textContent = payload.stop_phrase;
                    appendLog('system', 'Agent Initialized.');
                    break;
                case 'log':
                    appendLog('system', payload);
                    addWidgetLog(payload);
                    break;
                case 'status':
                    if (currentTask) currentTask.textContent = payload;
                    const s = (payload || '').toLowerCase();
                    if (s.includes('listening')) setWidgetState('listening');
                    else if (s.includes('process')) setWidgetState('processing');
                    else if (s.includes('speak') || s.includes('execut')) setWidgetState('speaking');
                    else setWidgetState('sleeping');
                    break;
                case 'mode_switch':
                    if (payload) {
                        if (micRing) micRing.className = 'mic-ring active';
                        if (modeState) {
                            modeState.textContent = 'Listening';
                            modeState.style.color = 'var(--accent-green)';
                        }
                        appendLog('system', 'WAKE WORD DETECTED - Listening...');
                        setWidgetState('listening');
                        addWidgetLog('Wake word detected');
                    } else {
                        if (micRing) micRing.className = 'mic-ring idle';
                        if (modeState) {
                            modeState.textContent = 'Idle';
                            modeState.style.color = 'var(--text-primary)';
                        }
                        appendLog('system', 'Stop phrase detected - Idle.');
                        setWidgetState('sleeping');
                    }
                    break;
                case 'transcription':
                    if (payload.text) {
                        if (payload.active && lastTranscription) {
                            lastTranscription.textContent = payload.text;
                        }
                        appendLog('transcription', `🗣️ "${payload.text}"`);
                        if (widgetTranscriptText) widgetTranscriptText.innerHTML = `"${payload.text}"`;
                    }
                    break;
                case 'decision':
                    appendLog('decision', `🧠 Decided Action: ${payload.action}`);
                    addWidgetLog(`Action: ${payload.action}`);
                    break;
                case 'action_result':
                    if (lastAction) lastAction.textContent = payload;
                    appendLog('action', `✅ Executed: ${payload}`);
                    addWidgetLog(`Done: ${payload}`);
                    break;
            }
        } catch(err) {
            console.error("Error parsing SSE event", err);
        }
    };

    evtSource.onerror = function(err) {
        console.error("EventSource failed:", err);
        appendLog('system', 'Lost connection to backend. Retrying...');
    };

    // --- Audio Control Logic ---
    const volSlider = document.getElementById('volume-slider');
    const volVal = document.getElementById('vol-val');
    const micSlider = document.getElementById('mic-slider');
    const micVal = document.getElementById('mic-val');
    const muteBtn = document.getElementById('mute-btn');

    let isMuted = false;
    let updateTimeout;

    function debounceUpdate(endpoint, payload) {
        clearTimeout(updateTimeout);
        updateTimeout = setTimeout(() => {
            fetch(endpoint, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            }).catch(err => console.error(`Failed to update ${endpoint}:`, err));
        }, 100);
    }

    if (volSlider) {
        volSlider.addEventListener('input', (e) => {
            const val = e.target.value;
            if (volVal) volVal.textContent = `${val}%`;
            const percentage = val + '%';
            volSlider.style.background = `linear-gradient(90deg, var(--accent-cyan) ${percentage}, rgba(0, 191, 255, 0.2) ${percentage})`;
            debounceUpdate('/set-volume', { volume: parseInt(val) });
        });
        volSlider.dispatchEvent(new Event('input'));
    }

    if (micSlider) {
        micSlider.addEventListener('input', (e) => {
            const val = e.target.value;
            if (micVal) micVal.textContent = `${val}%`;
            const percentage = val + '%';
            micSlider.style.background = `linear-gradient(90deg, var(--accent-orange) ${percentage}, rgba(0, 191, 255, 0.2) ${percentage})`;
            debounceUpdate('/set-mic', { sensitivity: parseInt(val) });
        });
        micSlider.dispatchEvent(new Event('input'));
    }

    if (muteBtn) {
        muteBtn.addEventListener('click', () => {
            isMuted = !isMuted;
            muteBtn.textContent = isMuted ? '🔇' : '🔊';
            muteBtn.classList.toggle('muted', isMuted);
            
            fetch('/toggle-mute', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ muted: isMuted })
            }).catch(err => console.error("Failed to toggle mute:", err));
        });
    }

    // --- Morphing Blob Canvas ---
    const canvas = document.getElementById('blobCanvas');
    if (canvas) {
        const ctx = canvas.getContext('2d');
        const W = 210, H = 210, CX = W / 2, CY = H / 2;
        let animFrame = 0;
        let volumeLevel = 0;

        function noise(t, freq, amp, phase) {
            return Math.sin(t * freq + phase) * amp;
        }

        function getBlobRadius(angle, t) {
            const base = 62;
            const stateAmps = {
                sleeping: [6, 3, 2, 1],
                listening: [12, 7, 5, 3],
                processing: [16, 10, 7, 4],
                speaking: [14, 8, 6, 3],
            };
            const amps = stateAmps[agentState] || stateAmps.sleeping;
            const speeds = {
                sleeping: 0.4,
                listening: 1.4,
                processing: 2.2,
                speaking: 1.8,
            };
            const speed = speeds[agentState] || 0.4;
            const vol = volumeLevel * 18;

            return (
                base +
                noise(angle, 2, amps[0] + vol, t * speed * 0.7) +
                noise(angle, 3, amps[1] + vol * 0.5, t * speed * 1.1 + 1.2) +
                noise(angle, 5, amps[2], t * speed * 1.7 + 2.4) +
                noise(angle, 7, amps[3], t * speed * 2.3 + 3.7)
            );
        }

        function getGradientColors() {
            const t = animFrame * 0.008;
            const r1 = Math.floor(80 + 40 * Math.sin(t * 0.7));
            const g1 = Math.floor(50 + 30 * Math.sin(t * 0.5 + 1));
            const b1 = Math.floor(200 + 55 * Math.sin(t * 0.9 + 2));
            const r2 = Math.floor(160 + 60 * Math.sin(t * 0.6 + 3));
            const g2 = Math.floor(30 + 20 * Math.sin(t * 0.8 + 0.5));
            const b2 = Math.floor(200 + 55 * Math.sin(t * 1.1));
            const r3 = Math.floor(40 + 40 * Math.sin(t * 0.4 + 1.5));
            const g3 = Math.floor(180 + 50 * Math.sin(t * 0.6 + 2));
            const b3 = Math.floor(230 + 25 * Math.sin(t * 0.7 + 3.5));
            return {
                c1: `rgb(${r1},${g1},${b1})`,
                c2: `rgb(${r2},${g2},${b2})`,
                c3: `rgb(${r3},${g3},${b3})`,
            };
        }

        function drawBlob(t) {
            ctx.clearRect(0, 0, W, H);
            const steps = 180;
            const pts = [];
            for (let i = 0; i <= steps; i++) {
                const angle = (i / steps) * Math.PI * 2;
                const r = getBlobRadius(angle, t);
                pts.push({ x: CX + r * Math.cos(angle), y: CY + r * Math.sin(angle) });
            }

            for (let g = 3; g >= 1; g--) {
                ctx.save();
                ctx.filter = `blur(${g * 7}px)`;
                ctx.globalAlpha = 0.18 / g;
                const colors = getGradientColors();
                const grad = ctx.createRadialGradient(CX - 10, CY - 15, 10, CX, CY, 78);
                grad.addColorStop(0, colors.c1);
                grad.addColorStop(0.5, colors.c2);
                grad.addColorStop(1, colors.c3);
                ctx.fillStyle = grad;
                ctx.beginPath();
                ctx.moveTo(pts[0].x, pts[0].y);
                for (let i = 1; i < pts.length; i++) {
                    const p = pts[i], pp = pts[i - 1];
                    ctx.quadraticCurveTo(pp.x, pp.y, (pp.x + p.x) / 2, (pp.y + p.y) / 2);
                }
                ctx.closePath();
                ctx.fill();
                ctx.restore();
            }

            const colors = getGradientColors();
            const grad = ctx.createRadialGradient(CX - 18, CY - 22, 8, CX + 10, CY + 10, 82);
            grad.addColorStop(0, colors.c1);
            grad.addColorStop(0.38, colors.c2);
            grad.addColorStop(0.72, `rgb(100,50,220)`);
            grad.addColorStop(1, colors.c3);
            ctx.globalAlpha = 0.92;
            ctx.fillStyle = grad;
            ctx.beginPath();
            ctx.moveTo(pts[0].x, pts[0].y);
            for (let i = 1; i < pts.length; i++) {
                const p = pts[i], pp = pts[i - 1];
                ctx.quadraticCurveTo(pp.x, pp.y, (pp.x + p.x) / 2, (pp.y + p.y) / 2);
            }
            ctx.closePath();
            ctx.fill();

            ctx.globalAlpha = 0.22;
            const sheen = ctx.createRadialGradient(CX - 22, CY - 26, 2, CX - 10, CY - 10, 55);
            sheen.addColorStop(0, 'rgba(255,255,255,0.9)');
            sheen.addColorStop(1, 'rgba(255,255,255,0)');
            ctx.fillStyle = sheen;
            ctx.beginPath();
            ctx.moveTo(pts[0].x, pts[0].y);
            for (let i = 1; i < pts.length; i++) {
                const p = pts[i], pp = pts[i - 1];
                ctx.quadraticCurveTo(pp.x, pp.y, (pp.x + p.x) / 2, (pp.y + p.y) / 2);
            }
            ctx.closePath();
            ctx.fill();

            ctx.globalAlpha = 0.08;
            for (let ring = 0; ring < 3; ring++) {
                const scale = 0.7 - ring * 0.13;
                ctx.save();
                ctx.translate(CX, CY);
                ctx.scale(scale, scale);
                ctx.translate(-CX, -CY);
                ctx.strokeStyle = 'rgba(200,180,255,0.8)';
                ctx.lineWidth = 1;
                ctx.beginPath();
                ctx.moveTo(pts[0].x, pts[0].y);
                for (let i = 1; i < pts.length; i++) {
                    const p = pts[i], pp = pts[i - 1];
                    ctx.quadraticCurveTo(pp.x, pp.y, (pp.x + p.x) / 2, (pp.y + p.y) / 2);
                }
                ctx.closePath();
                ctx.stroke();
                ctx.restore();
            }
            ctx.globalAlpha = 1;
        }

        function animate() {
            animFrame++;
            const t = animFrame * 0.016;
            drawBlob(t);
            requestAnimationFrame(animate);
        }
        animate();
    }

    // --- Drag Logic ---
    const widget = document.getElementById('widget-view');
    let dragging = false;
    let dragStartX = 0, dragStartY = 0;
    let widgetStartLeft = 0, widgetStartTop = 0;

    if (widget) {
        widget.addEventListener('mousedown', e => {
            // Allow dragging the mini assistant, but not other buttons
            const isMini = e.target.closest('.mini-assistant');
            if (e.target.closest('button') && !isMini) return;
            if (e.target.closest('.mic-icon')) return;
            
            dragging = true;
            hasDraggedWidget = false;
            dragStartX = e.screenX;
            dragStartY = e.screenY;
            const rect = widget.getBoundingClientRect();
            widgetStartLeft = rect.left;
            widgetStartTop = rect.top;
            widget.classList.add('dragging');
            
            // Prevent default to avoid text selection while dragging
            if (!isMini) {
                e.preventDefault();
            }
        });
        
        document.addEventListener('mousemove', e => {
            if (!dragging) return;
            const dx = e.screenX - dragStartX;
            const dy = e.screenY - dragStartY;
            
            if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
                hasDraggedWidget = true;
            }
            
            widget.style.left = `${Math.max(0, widgetStartLeft + dx)}px`;
            widget.style.top = `${Math.max(0, widgetStartTop + dy)}px`;
            widget.style.right = 'auto';
        });
        
        document.addEventListener('mouseup', () => {
            dragging = false;
            widget.classList.remove('dragging');
        });
    }
});
