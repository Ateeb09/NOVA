document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const micRing = document.getElementById('mic-ring');
    const modeState = document.getElementById('mode-state');
    const currentTask = document.getElementById('current-task');
    const wakeWordEl = document.getElementById('wake-word');
    const stopPhraseEl = document.getElementById('stop-phrase');
    const lastTranscription = document.getElementById('last-transcription');
    const lastAction = document.getElementById('last-action');
    const logContainer = document.getElementById('log-container');

    // Setup SSE Connection
    const evtSource = new EventSource("/events");

    function appendLog(type, message) {
        const div = document.createElement('div');
        div.className = `log-entry ${type}`;
        
        const now = new Date();
        const timeStr = now.toLocaleTimeString([], { hour12: false });
        
        div.innerHTML = `<span class="timestamp">[${timeStr}]</span> ${message}`;
        logContainer.appendChild(div);
        
        // Auto-scroll to bottom
        logContainer.scrollTop = logContainer.scrollHeight;
    }

    evtSource.onmessage = function(event) {
        try {
            const data = JSON.parse(event.data);
            const type = data.type;
            const payload = data.payload;

            switch(type) {
                case 'init':
                    wakeWordEl.textContent = payload.wake_word;
                    stopPhraseEl.textContent = payload.stop_phrase;
                    appendLog('system', 'Agent Initialized.');
                    break;
                case 'log':
                    appendLog('system', payload);
                    break;
                case 'status':
                    currentTask.textContent = payload;
                    break;
                case 'mode_switch':
                    // Payload is boolean (true = active, false = idle)
                    if (payload) {
                        micRing.className = 'mic-ring active';
                        modeState.textContent = 'Listening';
                        modeState.style.color = 'var(--accent-green)';
                        appendLog('system', 'WAKE WORD DETECTED - Listening...');
                    } else {
                        micRing.className = 'mic-ring idle';
                        modeState.textContent = 'Idle';
                        modeState.style.color = 'var(--text-primary)';
                        appendLog('system', 'Stop phrase detected - Idle.');
                    }
                    break;
                case 'transcription':
                    // Payload: { text: "...", active: bool }
                    if (payload.text) {
                        if (payload.active) {
                            lastTranscription.textContent = payload.text;
                        }
                        appendLog('transcription', `🗣️ "${payload.text}"`);
                    }
                    break;
                case 'decision':
                    // Payload: { command: "...", action: "..." }
                    appendLog('decision', `🧠 Decided Action: ${payload.action}`);
                    break;
                case 'action_result':
                    lastAction.textContent = payload;
                    appendLog('action', `✅ Executed: ${payload}`);
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
        }, 100); // 100ms debounce prevents lagging the backend
    }

    if (volSlider) {
        volSlider.addEventListener('input', (e) => {
            const val = e.target.value;
            volVal.textContent = `${val}%`;
            
            // Visual fill effect on slider background
            const percentage = val + '%';
            volSlider.style.background = `linear-gradient(90deg, var(--accent-cyan) ${percentage}, rgba(0, 191, 255, 0.2) ${percentage})`;
            
            debounceUpdate('/set-volume', { volume: parseInt(val) });
        });
        // Initialize gradient background
        volSlider.dispatchEvent(new Event('input'));
    }

    if (micSlider) {
        micSlider.addEventListener('input', (e) => {
            const val = e.target.value;
            micVal.textContent = `${val}%`;
            
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
});
