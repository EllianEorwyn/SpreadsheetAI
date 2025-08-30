/*
 * Spreadsheet AI Processor — Front‑end logic (app.js)
 * ---------------------------------------------------
 * This script powers the browser‑based interface for running large‑language‑model
 * (LLM) analyses over spreadsheet rows. It handles:
 *   • Caching DOM element references
 *   • Centralised application state management
 *   • Dynamic task creation & removal (analyze / compare / custom / auto)
 *   • CSV parsing & data preview rendering
 *   • Prompt construction and provider‑agnostic API calls (OpenAI, Gemini, Ollama)
 *   • Cost / token estimation
 *   • Result validation utilities (reliability, consistency, face‑validity, etc.)
 *   • Inline dashboard visualisations (Chart.js & WordCloud)
 *   • Test‑mode stepping, progress UI, and audit logging
 *
 * The codebase is divided into clear sections.  Each section is now preceded by
 * a short JSDoc‑style block explaining its role to help new contributors get
 * oriented quickly.
 */

document.addEventListener('DOMContentLoaded', () => {

    // --- DOM Elements ---
/*
 * This block caches frequent DOM look‑ups in a single object literal (`ui`).
 * By keeping a stable reference, we avoid repeated `document.getElementById` calls
 * and make component access uniform across the file.
 */
    const ui = {
        fileInput: document.getElementById('file-input'),
        fileInfo: document.getElementById('file-info'),
        dataPreview: document.getElementById('data-preview'),
        auditLog: document.getElementById('audit-log'),
        statusIndicator: document.getElementById('status-indicator'),
        controlsPanel: document.getElementById('controls-panel'),
        providerSelector: document.getElementById('provider-selector'),
        geminiConfig: document.getElementById('gemini-config'),
        openaiConfig: document.getElementById('openai-config'),
        ollamaConfig: document.getElementById('ollama-config'),
        geminiApiKeyInput: document.getElementById('gemini-api-key-input'),
        geminiModelSelector: document.getElementById('gemini-model-selector'),
        openaiApiKeyInput: document.getElementById('openai-api-key-input'),
        fetchOpenAIModelsBtn: document.getElementById('fetch-openai-models-btn'),
        openaiModelSelector: document.getElementById('openai-model-selector'),
        openaiParamsBtn: document.getElementById('openai-params-btn'),
        geminiParamsBtn: document.getElementById('gemini-params-btn'),
        ollamaParamsBtn: document.getElementById('ollama-params-btn'),
        aiParamsPanel: document.getElementById('ai-params-panel'),
        cachedInputsToggle: document.getElementById('cached-inputs-toggle'),
        temperatureSlider: document.getElementById('temperature-slider'),
        maxTokensSlider: document.getElementById('max-tokens-slider'),
        topPSlider: document.getElementById('top-p-slider'),
        nSlider: document.getElementById('n-slider'),
        stopSlider: document.getElementById('stop-slider'),
        freqPenaltySlider: document.getElementById('freq-penalty-slider'),
        presPenaltySlider: document.getElementById('pres-penalty-slider'),
        temperatureValue: document.getElementById('temperature-value'),
        maxTokensValue: document.getElementById('max-tokens-value'),
        topPValue: document.getElementById('top-p-value'),
        nValue: document.getElementById('n-value'),
        stopValue: document.getElementById('stop-value'),
        freqPenaltyValue: document.getElementById('freq-penalty-value'),
        presPenaltyValue: document.getElementById('pres-penalty-value'),
        gpt5Options: document.getElementById('gpt5-extra-params'),
        gpt5ReasoningSelect: document.getElementById('gpt5-reasoning'),
        gpt5VerbositySelect: document.getElementById('gpt5-verbosity'),
        ollamaOptions: document.getElementById('ollama-extra-params'),
        ollamaReasoningToggle: document.getElementById('ollama-reasoning-toggle'),
        ollamaUrlInput: document.getElementById('ollama-url-input'),
        fetchOllamaModelsBtn: document.getElementById('fetch-ollama-models-btn'),
        ollamaModelSelector: document.getElementById('ollama-model-selector'),
        systemPromptInput: document.getElementById('system-prompt-input'),
        projectDescriptionInput: document.getElementById('project-description-input'),
        additionalContextInput: document.getElementById('additional-context-input'),
        guessProjectBtn: document.getElementById('guess-project-btn'),
        guessAdditionalBtn: document.getElementById('guess-additional-btn'),
        addTaskDropdownContainer: document.getElementById('add-task-dropdown-container'),
        addTaskBtn: document.getElementById('add-task-btn'),
        addTaskMenu: document.getElementById('add-task-menu'),
        addPresetDropdownContainer: document.getElementById('add-preset-dropdown-container'),
        addPresetBtn: document.getElementById('add-preset-btn'),
        addPresetMenu: document.getElementById('add-preset-menu'),
        taskContainer: document.getElementById('task-container'),
        analyzeTaskTemplate: document.getElementById('analyze-task-template'),
        compareTaskTemplate: document.getElementById('compare-task-template'),
        customTaskTemplate: document.getElementById('custom-task-template'),
        autoTaskTemplate: document.getElementById("auto-task-template"),
        costEstimate: document.getElementById('cost-estimate'),
        tokenEstimate: document.getElementById('token-estimate'),
        runBtn: document.getElementById('run-btn'),
        testModeBtn: document.getElementById('test-mode-btn'),
        progressBarContainer: document.getElementById('progress-bar-container'),
        progressBar: document.getElementById('progress-bar'),
        downloadResultsBtn: document.getElementById('download-results-btn'),
        downloadLogBtn: document.getElementById('download-log-btn'),
        selectTaskModal: document.getElementById('select-task-modal'),
        closeSelectTaskModalBtn: document.getElementById('close-select-task-modal-btn'),
        testableTasksList: document.getElementById('testable-tasks-list'),
        testModeModal: document.getElementById('test-mode-modal'),
        closeTestModeBtn: document.getElementById('close-test-mode-btn'),
        testModeTitle: document.getElementById('test-mode-title'),
        testModeNavInfo: document.getElementById('test-mode-nav-info'),
        testModeRowData: document.getElementById('test-mode-row-data'),
        testModePrompt: document.getElementById('test-mode-prompt'),
        testModeOutput: document.getElementById('test-mode-output'),
        testModePrevBtn: document.getElementById('test-mode-prev-btn'),
        testModeRerunBtn: document.getElementById('test-mode-rerun-btn'),
        testModeNextBtn: document.getElementById('test-mode-next-btn'),
        saveProfileBtn: document.getElementById('save-profile-btn'),
        loadProfileInput: document.getElementById('load-profile-input'),
        presetModal: document.getElementById('preset-modal'),
        presetModalTitle: document.getElementById('preset-modal-title'),
        presetModalBody: document.getElementById('preset-modal-body'),
        confirmPresetBtn: document.getElementById('confirm-preset-btn'),
        cancelPresetBtn: document.getElementById('cancel-preset-btn'),
        includeRowContextToggle: document.getElementById('include-row-context-toggle'),
        checkMissingOutputs: document.getElementById('check-missing-outputs'),
        promptStability: document.getElementById('prompt-stability'),
        consistencyCheck: document.getElementById('consistency-check'),
        consistencyColumns: document.getElementById('consistency-columns'),
        faceValidity: document.getElementById('face-validity'),
        analysisDashboardPanel: document.getElementById('analysis-dashboard-panel'),
        analysisDashboard: document.getElementById('analysis-dashboard'),
        rightPanel: document.getElementById('right-panel'),
    };

    // --- State Management ---
/*
 * All ephemeral data that drives the UI lives in `appState`. Treat it as a
 * single‑source‑of‑truth; UI setters & helper functions read/write here so that the
 * interface stays in sync.
 */
    let appState = {
        file: null,
        data: [],
        headers: [],
        isProcessing: false,
        runLog: {
            startTime: new Date().toISOString(),
            entries: [],
            totalInputTokens: 0,
            totalOutputTokens: 0,
            validations: [],
            aiParams: {}
        },
        analysisTasks: [],
        availableModels: { openai: [], ollama: [] },
        testMode: { isActive: false, task: null, currentIndex: 0 },
        includeRowContext: true,
        validationSettings: { consistencyCheck: false, consistencyColumns: [], missingCheck: false, faceValidity: false, promptStability: false },
        currentPreset: null,
        dashboardOverrides: {},
        aiParams: {
            cache: false,
            temperature: 1,
            maxTokens: 150,
            topP: 1,
            n: 1,
            stop: 0,
            freqPenalty: 0,
            presPenalty: 0,
            reasoning: 'minimal',
            verbosity: 'medium',
            ollamaReasoning: false
        }
    };
    
    const MODEL_PRICING = {
        'gpt-4.1': { input: 2.00, cached: 0.50, output: 8.00 },
        'gpt-4.1-mini': { input: 0.40, cached: 0.10, output: 1.60 },
        'gpt-4.1-nano': { input: 0.10, cached: 0.03, output: 0.40 },
        'gpt-4.5-preview': { input: 75.00, cached: 37.50, output: 150.00 },
        'gpt-4o': { input: 2.50, cached: 1.25, output: 10.00 },
        'gpt-4o-audio-preview': { input: 2.50, cached: 0, output: 10.00 },
        'gpt-4o-realtime-preview': { input: 5.00, cached: 2.50, output: 20.00 },
        'gpt-4o-mini': { input: 0.15, cached: 0.08, output: 0.60 },
        'gpt-4o-mini-audio-preview': { input: 0.15, cached: 0, output: 0.60 },
        'gpt-4o-mini-realtime-preview': { input: 0.60, cached: 0.30, output: 2.40 },
        'o1': { input: 15.00, cached: 7.50, output: 60.00 },
        'o1-pro': { input: 150.00, cached: 0, output: 600.00 },
        'o3-pro': { input: 20.00, cached: 0, output: 80.00 },
        'o3': { input: 2.00, cached: 0.50, output: 8.00 },
        'o3-deep-research': { input: 10.00, cached: 2.50, output: 40.00 },
        'o4-mini': { input: 1.10, cached: 0.28, output: 4.40 },
        'o4-mini-deep-research': { input: 2.00, cached: 0.50, output: 8.00 },
        'o3-mini': { input: 1.10, cached: 0.55, output: 4.40 },
        'o1-mini': { input: 1.10, cached: 0.55, output: 4.40 },
        'gemini-1.5-flash-latest': { input: 0.35, cached: 0, output: 0.70 },
        'default_openai': { input: 1.00, cached: 0.00, output: 3.00 }
    };
    const DEFAULT_SYSTEM_PROMPT = "You are an AI assistant whose sole task is to process spreadsheet rows exactly as the user’s analysis tasks specify. Output only the requested content in the exact format required—no greetings, acknowledgments, or extra commentary.";

    const TASK_PRESETS = {
      "sentiment": {
        label: "Sentiment Classification",
        type: "analyze",
        promptTemplate: "Classify the sentiment of this text as Positive, Negative, or Neutral:\n\n{{COLUMN}}",
        outputColumn: "sentiment"
      },
      "summarize": {
        label: "Summarize Text",
        type: "analyze",
        promptTemplate: "Summarize the main idea of the following text in 1-2 sentences:\n\n{{COLUMN}}",
        outputColumn: "summary"
      },
      "topic": {
        label: "Topic Detection",
        type: "analyze",
        promptTemplate: "Identify the topic of this text in one or two words (e.g., climate, finance, health):\n\n{{COLUMN}}",
        outputColumn: "topic"
      },
      "emotion": {
        label: "Emotion Classification",
        type: "analyze",
        promptTemplate: "Identify the dominant emotion expressed (e.g., joy, anger, sadness, fear, surprise):\n\n{{COLUMN}}",
        outputColumn: "emotion"
      },
      "stance": {
        label: "Identify Stance",
        type: "compare",
        promptTemplate: "Compare the following two statements and determine whether the second one supports, opposes, or is neutral toward the first:\n\n{{COLUMNS_DATA}}",
        outputColumn: "stance"
      },
      "contradiction": {
        label: "Detect Contradictions",
        type: "compare",
        promptTemplate: "Determine whether these two statements contradict each other:\n\n{{COLUMNS_DATA}}",
        outputColumn: "contradiction"
      },
      "headline": {
        label: "Generate Headline",
        type: "custom",
        promptTemplate: "Write a short, engaging headline for the following row:\n\n{{Column1}} {{Column2}} {{Column3}}",
        outputColumn: "headline"
      },
      "entities": {
        label: "Extract Entities",
        type: "custom",
        promptTemplate: "List all named people, organizations, and places mentioned in the text:\n\n{{text}}",
        outputColumn: "entities"
      }
    };

    // --- Function Definitions ---
/*
 * Helper functions are defined next.  They are organised roughly from generic
 * utilities (logging, downloads, cost estimation) to higher‑level operations
 * (prompt building, validation checks, API wrappers).
 */

    const log = (message, type = 'SYSTEM') => {
        const p = document.createElement('p');
        p.innerHTML = `<span class="text-gray-500">${new Date().toLocaleTimeString()} [${type}]</span> ${message}`;
        if (type === 'ERROR') p.classList.add('text-red-400');
        ui.auditLog.appendChild(p);
        ui.auditLog.scrollTop = ui.auditLog.scrollHeight;
    };

    const estimateTokens = (text) => Math.ceil((text || '').length / 4);

    const parseCodeSet = (str) => {
        if (str === undefined || str === null) return [];
        return String(str).split(/[;,]/).map(s => s.trim()).filter(Boolean);
    };

    const downloadFile = (content, fileName, mimeType) => {
        const a = document.createElement('a');
        const blob = new Blob([content], { type: mimeType });
        a.href = URL.createObjectURL(blob);
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(a.href);
    };

    const recordAuditEntry = (entry) => {
        appState.runLog.entries.push(entry);
        appState.runLog.totalInputTokens += entry.inputTokens || 0;
        appState.runLog.totalOutputTokens += entry.outputTokens || 0;
        ui.downloadLogBtn.disabled = false;
        log(`Logged prompt (${entry.inputTokens} in, ${entry.outputTokens} out tokens).`, 'AUDIT');
    };

    const setProcessingState = (isProcessing) => {
        appState.isProcessing = isProcessing;
        ui.controlsPanel.classList.toggle('disabled-ui', isProcessing);
        const statusSvg = isProcessing ? `<svg class="w-4 h-4 mr-2 animate-spin text-blue-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 6v6m0 0v6m0-6h6m-6 0H6"></path></svg>` : `<svg class="w-4 h-4 mr-2 text-green-400" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm-1-11a1 1 0 10-2 0v4a1 1 0 102 0V7zm2-1a1 1 0 00-1 1v4a1 1 0 102 0V7a1 1 0 00-1-1z" clip-rule="evenodd"></path></svg>`;
        ui.statusIndicator.innerHTML = `${statusSvg}${isProcessing ? 'Processing...' : 'Ready'}`;
        ui.progressBarContainer.style.display = isProcessing ? 'block' : 'hidden';
        ui.progressBar.style.width = '0%';
    };

    const renderDataPreview = () => {
        if (appState.data.length === 0) {
            ui.dataPreview.innerHTML = `<p class="p-4 text-gray-500">Upload a CSV file to see a preview of the data here.</p>`;
            return;
        }
        const table = document.createElement('table');
        table.className = 'w-full text-sm text-left text-gray-400';
        const thead = document.createElement('thead');
        thead.className = 'text-xs text-gray-300 uppercase bg-gray-700 sticky top-0';
        let tr = document.createElement('tr');
        appState.headers.forEach(header => {
            const th = document.createElement('th');
            th.scope = 'col';
            th.className = 'px-4 py-2';
            th.textContent = header;
            tr.appendChild(th);
        });
        thead.appendChild(tr);
        table.appendChild(thead);
        const tbody = document.createElement('tbody');
        appState.data.slice(0, 100).forEach(row => {
            tr = document.createElement('tr');
            tr.className = 'bg-gray-800 border-b border-gray-700';
            appState.headers.forEach(header => {
                const td = document.createElement('td');
                td.className = 'px-4 py-2 whitespace-nowrap overflow-hidden text-ellipsis max-w-xs';
                td.textContent = row[header];
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        ui.dataPreview.innerHTML = '';
        ui.dataPreview.appendChild(table);
    };

    const detectColumnInfo = (values) => {
        const clean = values.filter(v => v !== undefined && v !== null && v !== '' && v !== 'ERROR');
        const blankPct = values.length ? ((values.length - clean.length) / values.length * 100).toFixed(1) : 0;
        const numericVals = clean.map(v => parseFloat(v)).filter(v => !isNaN(v));
        let type = 'text';
        let codes = [];
        if (numericVals.length === clean.length && clean.length > 0) {
            type = 'numeric';
        } else {
            const unique = [...new Set(clean)];
            const avgLen = clean.reduce((a, b) => a + String(b).length, 0) / (clean.length || 1);
            if (unique.length <= 20 && avgLen <= 20) {
                type = 'categorical';
            } else {
                const splitCodes = clean.flatMap(v => parseCodeSet(v));
                const uniqueCodes = [...new Set(splitCodes)];
                if (uniqueCodes.length > 0 && uniqueCodes.length <= 20) {
                    type = 'list';
                    codes = uniqueCodes;
                }
            }
        }
        return { clean, blankPct, type, numericVals, codes };
    };

    const createFrequencyTable = (counts, total) => {
        const table = document.createElement('table');
        table.className = 'w-full text-xs text-left mt-2';
        const thead = document.createElement('thead');
        thead.innerHTML = '<tr><th class="pr-2">Value</th><th class="pr-2">Count</th><th>%</th></tr>';
        table.appendChild(thead);
        const tbody = document.createElement('tbody');
        Object.entries(counts).sort((a,b) => b[1]-a[1]).forEach(([val,c]) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td class="pr-2">${val}</td><td class="pr-2">${c}</td><td>${((c/total)*100).toFixed(1)}</td>`;
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        return table;
    };

    const createNumericStats = (nums) => {
        const min = Math.min(...nums);
        const max = Math.max(...nums);
        const mean = nums.reduce((a,b) => a+b,0) / nums.length;
        const sd = Math.sqrt(nums.reduce((a,b)=>a+Math.pow(b-mean,2),0)/nums.length);
        const outliers = nums.filter(v => Math.abs(v-mean) > 2*sd).length;
        const div = document.createElement('div');
        div.className = 'text-xs mt-2 text-gray-300';
        div.textContent = `Min: ${min.toFixed(2)} Max: ${max.toFixed(2)} Mean: ${mean.toFixed(2)} StdDev: ${sd.toFixed(2)} Outliers: ${outliers}`;
        return div;
    };

    const createBarChart = (labels, data, logScale=false) => {
        const canvas = document.createElement('canvas');
        new Chart(canvas, {
            type: 'bar',
            data: { labels, datasets: [{ label: 'Count', data, backgroundColor: '#60a5fa' }] },
            options: { scales: { y: { beginAtZero: true, type: logScale ? 'logarithmic' : 'linear' } } }
        });
        return canvas;
    };

    const createHistogram = (nums, bins=10, logScale=false) => {
        const min = Math.min(...nums);
        const max = Math.max(...nums);
        const step = (max - min) / bins || 1;
        const counts = Array(bins).fill(0);
        nums.forEach(v => { const idx = Math.min(bins-1, Math.floor((v-min)/step)); counts[idx]++; });
        const labels = counts.map((_,i)=>`${(min+i*step).toFixed(1)}-${(min+(i+1)*step).toFixed(1)}`);
        return createBarChart(labels, counts, logScale);
    };

    const createWordCloud = (words) => {
        const canvas = document.createElement('canvas');
        const list = Object.entries(words).map(([w,c]) => [w, c]);
        setTimeout(() => WordCloud(canvas, { list, backgroundColor:'#111827' }), 0);
        return canvas;
    };

    const renderAnalysisDashboard = () => {
        const outputCols = appState.analysisTasks.map(t => t.outputColumn).filter(c => c && appState.headers.includes(c));
        if (outputCols.length === 0) { ui.analysisDashboardPanel.classList.add('hidden'); return; }
        ui.analysisDashboard.innerHTML = '';
        outputCols.forEach(col => {
            const values = appState.data.map(r => r[col]);
            const info = detectColumnInfo(values);
            const override = appState.dashboardOverrides[col];
            const chosenType = override || info.type;

            const wrapper = document.createElement('div');
            wrapper.className = 'border border-gray-700 rounded';

            const header = document.createElement('div');
            header.className = 'w-full px-2 py-1 bg-gray-700 flex justify-between items-center cursor-pointer';
            const left = document.createElement('div');
            left.className = 'flex items-center space-x-2';
            const nameSpan = document.createElement('span');
            nameSpan.textContent = col;
            const typeSelect = document.createElement('select');
            ['auto','categorical','text'].forEach(opt => {
                const o = document.createElement('option');
                o.value = opt;
                o.textContent = opt.charAt(0).toUpperCase() + opt.slice(1);
                typeSelect.appendChild(o);
            });
            typeSelect.value = override || 'auto';
            typeSelect.className = 'bg-gray-600 text-xs text-white rounded';
            typeSelect.addEventListener('change', (e) => {
                const val = e.target.value;
                if (val === 'auto') delete appState.dashboardOverrides[col];
                else appState.dashboardOverrides[col] = val;
                renderAnalysisDashboard();
                e.stopPropagation();
            });
            left.appendChild(nameSpan);
            left.appendChild(typeSelect);
            const toggle = document.createElement('span');
            toggle.className = 'toggle';
            toggle.textContent = '+';
            header.appendChild(left);
            header.appendChild(toggle);

            const content = document.createElement('div');
            content.className = 'hidden p-2';
            const summary = document.createElement('div');
            summary.className = 'text-sm mb-2 text-gray-300';
            summary.textContent = `Non-null: ${info.clean.length} | Blank/Error: ${info.blankPct}% | Type: ${chosenType}`;
            content.appendChild(summary);

            let chart, table, stats;
            if (chosenType === 'numeric') {
                chart = createHistogram(info.numericVals);
                stats = createNumericStats(info.numericVals);
                content.appendChild(chart);
                content.appendChild(stats);
            } else if (chosenType === 'categorical') {
                const counts = {};
                info.clean.forEach(v => { counts[v] = (counts[v]||0)+1; });
                chart = createBarChart(Object.keys(counts), Object.values(counts));
                table = createFrequencyTable(counts, info.clean.length);
                content.appendChild(chart);
                content.appendChild(table);
            } else if (chosenType === 'list') {
                const counts = {};
                info.clean.forEach(v => parseCodeSet(v).forEach(code => { counts[code] = (counts[code]||0)+1; }));
                chart = createBarChart(Object.keys(counts), Object.values(counts));
                table = createFrequencyTable(counts, info.clean.length);
                content.appendChild(chart);
                content.appendChild(table);
            } else {
                const words = {};
                info.clean.forEach(v => String(v).toLowerCase().split(/\W+/).forEach(w => { if(w) words[w]=(words[w]||0)+1; }));
                chart = createWordCloud(words);
                table = createFrequencyTable(words, Object.values(words).reduce((a,b)=>a+b,0));
                content.appendChild(chart);
                content.appendChild(table);
            }

            const actions = document.createElement('div');
            actions.className = 'text-xs mt-2 space-x-2';
            if (chart) {
                const dl = document.createElement('button');
                dl.className = 'bg-gray-600 hover:bg-gray-500 text-white py-1 px-2 rounded';
                dl.textContent = 'Download PNG';
                dl.onclick = () => {
                    const link = document.createElement('a');
                    link.href = chart.toDataURL ? chart.toDataURL('image/png') : chart.firstChild.toDataURL('image/png');
                    link.download = `${col}.png`;
                    link.click();
                };
                actions.appendChild(dl);
            }
            if (table) {
                const cp = document.createElement('button');
                cp.className = 'bg-gray-600 hover:bg-gray-500 text-white py-1 px-2 rounded';
                cp.textContent = 'Copy Table';
                cp.onclick = () => {
                    const rows = Array.from(table.querySelectorAll('tbody tr')).map(r => Array.from(r.children).map(td => td.textContent).join(','));
                    navigator.clipboard.writeText(rows.join('\n'));
                };
                actions.appendChild(cp);
            }
            if (chosenType === 'numeric') {
                const logLabel = document.createElement('label');
                logLabel.className = 'ml-2';
                const chk = document.createElement('input');
                chk.type = 'checkbox';
                chk.className = 'mr-1';
                logLabel.appendChild(chk);
                logLabel.appendChild(document.createTextNode('Log scale'));
                chk.addEventListener('change', () => {
                    chart.config.options.scales.y.type = chk.checked ? 'logarithmic' : 'linear';
                    chart.update();
                });
                actions.appendChild(logLabel);
            }
            if (actions.childNodes.length) content.appendChild(actions);

            header.addEventListener('click', (e) => {
                if (e.target.tagName.toLowerCase() === 'select') return;
                content.classList.toggle('hidden');
            });

            wrapper.appendChild(header);
            wrapper.appendChild(content);
            ui.analysisDashboard.appendChild(wrapper);
        });
        ui.analysisDashboardPanel.classList.remove('hidden');
    };

    const renderReliabilityDashboard = () => {
        const panel = document.getElementById('reliability-panel');
        const container = document.getElementById('reliability-dashboard');
        if (!panel || !container) return;

        const results = (appState.runLog.validations || []).filter(v => v.validation === 'Intercoder Reliability');
        container.innerHTML = '';

        if (results.length === 0) { panel.classList.add('hidden'); return; }

        results.forEach(r => {
            const wrapper = document.createElement('div');
            wrapper.className = 'border border-gray-700 rounded';

            const header = document.createElement('div');
            header.className = 'w-full px-2 py-1 bg-gray-700 flex justify-between items-center cursor-pointer';
            const title = document.createElement('span');
            title.textContent = `${r.task} vs ${r.reference_column}`;
            const toggle = document.createElement('span');
            toggle.className = 'toggle';
            toggle.textContent = '+';
            header.appendChild(title);
            header.appendChild(toggle);

            const content = document.createElement('div');
            content.className = 'hidden p-2';

            const table = document.createElement('table');
            table.className = 'text-xs text-gray-300 mb-2';
            table.innerHTML = `<thead><tr><th class="px-2 py-1 text-left">Metric</th><th class="px-2 py-1">Score</th></tr></thead>` +
                `<tbody>` +
                `<tr><td class="px-2 py-1">Cohen's Kappa</td><td>${r.score}</td></tr>` +
                `<tr><td class="px-2 py-1">Krippendorff's Alpha (${r.distance})</td><td>${r.krippendorff_alpha}</td></tr>` +
                `<tr><td class="px-2 py-1">Observed Disagreement</td><td>${r.observed_disagreement}</td></tr>` +
                `<tr><td class="px-2 py-1">Expected Disagreement</td><td>${r.expected_disagreement}</td></tr>` +
                `<tr><td class="px-2 py-1">Exact Match</td><td>${r.exact_match}</td></tr>` +
                `</tbody>`;
            content.appendChild(table);

            const canvas = document.createElement('canvas');
            content.appendChild(canvas);
            new Chart(canvas, {
                type: 'bar',
                data: {
                    labels: ["Kappa", "Alpha", "Exact", "D_o", "D_e"],
                    datasets: [{
                        label: 'Score',
                        data: [r.score, r.krippendorff_alpha, r.exact_match, r.observed_disagreement, r.expected_disagreement],
                        backgroundColor: ['#60a5fa', '#c084fc', '#4ade80', '#f472b6', '#fbbf24']
                    }]
                },
                options: { scales: { y: { beginAtZero: true, max: 1 } } }
            });

            header.addEventListener('click', (e) => {
                if (e.target.tagName.toLowerCase() === 'canvas') return;
                content.classList.toggle('hidden');
            });

            wrapper.appendChild(header);
            wrapper.appendChild(content);
            container.appendChild(wrapper);
        });

        panel.classList.remove('hidden');
    };

    const getModelIdentifier = () => {
        const provider = ui.providerSelector.value;
        switch(provider) {
            case 'gemini': return ui.geminiModelSelector.value;
            case 'openai': return ui.openaiModelSelector.value;
            case 'ollama': return ui.ollamaModelSelector.value;
            default: return 'unknown';
        }
    };
    
    function buildRowContext(row) {
        const rowLines = appState.headers.map(h => `${h}: ${row[h] || ''}`);
        let context = '### ROW CONTEXT - DO NOT INCLUDE IN OUTPUT ###\n' + rowLines.join('\n');
        if (appState.data.length > 1 && appState.data[0] !== row) {
            const refLines = appState.headers.map(h => `${h}: ${appState.data[0][h] || ''}`);
            context += `\n\nReference Row:\n${refLines.join('\n')}`;
        }
        context += '\n### END CONTEXT ###';
        return context;
    }

    function buildPrompt(task, row) {
        let prompt = task.prompt;

        if (task.type === 'analyze') {
            prompt = prompt.replace(/\{\{COLUMN\}\}/g, row[task.sourceColumn] || '');
        } else if (task.type === 'compare') {
            const columnsData = task.sourceColumns.map(sc => `Column '${sc}': ${row[sc] || ''}`).join('\n');
            prompt = prompt.replace(/\{\{COLUMNS_DATA\}\}/g, columnsData);
        }

        appState.headers.forEach(header => {
            const escapedHeader = header.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
            const regex = new RegExp(`\\{\\{${escapedHeader}\\}\\}`, 'g');
            prompt = prompt.replace(regex, row[header] || '');
        });

        return `${appState.includeRowContext ? buildRowContext(row) + '\n\n' : ''}${prompt}`;
    }

    const updateCostEstimate = () => {
        if (!appState.data || appState.data.length === 0 || appState.analysisTasks.length === 0) {
            ui.costEstimate.textContent = "$0.00";
            ui.tokenEstimate.textContent = "~0 tokens";
            return;
        }
        
        const provider = ui.providerSelector.value;
        if (provider === 'ollama') {
            ui.costEstimate.textContent = "$0.00";
            ui.tokenEstimate.textContent = "Processing is local";
            return;
        }

        let totalInputTokens = 0;
        let totalOutputTokens = 0;
        const firstRow = appState.data[0];
        const systemPromptTokens = estimateTokens(ui.systemPromptInput.value + ui.projectDescriptionInput.value + ui.additionalContextInput.value);
        
        appState.analysisTasks.forEach(task => {
            if (task.prompt && firstRow) {
                const userPrompt = buildPrompt(task, firstRow);
                totalInputTokens += (estimateTokens(userPrompt) + systemPromptTokens) * appState.data.length;
                totalOutputTokens += (parseInt(task.maxTokens, 10) || 150) * appState.data.length;
            }
        });
        
        const completions = appState.aiParams.n || 1;
        const totalTokens = totalInputTokens + (totalOutputTokens * completions);
        const modelName = getModelIdentifier();
        const pricing = MODEL_PRICING[modelName] || MODEL_PRICING['default_openai'];
        const inputRate = (provider === 'openai' && appState.aiParams.cache) ? pricing.cached : pricing.input;
        const cost = ((totalInputTokens / 1_000_000) * inputRate) + (((totalOutputTokens * completions) / 1_000_000) * pricing.output);

        ui.costEstimate.textContent = `$${cost.toFixed(4)}`;
        ui.tokenEstimate.textContent = `~${totalTokens.toLocaleString()} tokens`;
    };

    function computeKappa(a, b) {
        const n = a.length;
        const cats = [...new Set([...a, ...b])];
        const freqA = {}, freqB = {};
        cats.forEach(c => { freqA[c] = 0; freqB[c] = 0; });
        let agree = 0;
        for (let i = 0; i < n; i++) {
            const ai = a[i];
            const hi = b[i];
            if (ai === hi) agree++;
            freqA[ai] = (freqA[ai] || 0) + 1;
            freqB[hi] = (freqB[hi] || 0) + 1;
        }
        const p0 = agree / n;
        let pe = 0;
        cats.forEach(c => { pe += (freqA[c] / n) * (freqB[c] / n); });
        return (p0 - pe) / (1 - pe);
    }

    function jaccardDistance(a, b) {
        const A = new Set(a);
        const B = new Set(b);
        if (A.size === 0 && B.size === 0) return 0;
        let intersect = 0;
        A.forEach(v => { if (B.has(v)) intersect++; });
        const union = new Set([...A, ...B]).size;
        return 1 - intersect / union;
    }

    function binaryDistance(a, b) {
        return canonicalizeSet(a) === canonicalizeSet(b) ? 0 : 1;
    }

    function masiDistance(a, b) {
        const A = new Set(a);
        const B = new Set(b);
        const interSize = [...A].filter(x => B.has(x)).length;
        const unionSize = new Set([...A, ...B]).size;
        if (unionSize === 0) return 0;
        if (interSize === 0) return 1;
        const A_subset_B = [...A].every(x => B.has(x));
        const B_subset_A = [...B].every(x => A.has(x));
        let w;
        if (A_subset_B && B_subset_A) w = 1;
        else if (A_subset_B || B_subset_A) w = 1/3;
        else w = 2/3;
        const similarity = w * interSize / unionSize;
        return 1 - similarity;
    }

    function canonicalizeSet(set) {
        return Array.from(new Set(set)).sort().join('|');
    }

    function decodeSet(str) {
        return str ? str.split('|') : [];
    }

    function computeKrippendorffAlpha(setsA, setsB, method = 'jaccard') {
        const n = setsA.length;
        if (n === 0) return { alpha: 0, Do: 0, De: 0 };

        const distanceFn = method === 'binary' ? binaryDistance : (method === 'masi' ? masiDistance : jaccardDistance);

        let Do = 0;
        for (let i = 0; i < n; i++) {
            Do += distanceFn(setsA[i], setsB[i]);
        }
        Do /= n;

        const freq = {};
        setsA.concat(setsB).forEach(set => {
            const key = canonicalizeSet(set);
            freq[key] = (freq[key] || 0) + 1;
        });
        const total = setsA.length + setsB.length;
        const keys = Object.keys(freq);

        let De = 0;
        for (const k1 of keys) {
            for (const k2 of keys) {
                const p = (freq[k1] / total) * (freq[k2] / total);
                if (p === 0) continue;
                De += p * distanceFn(decodeSet(k1), decodeSet(k2));
            }
        }

        if (De === 0) return { alpha: Do === 0 ? 1 : 0, Do, De };
        return { alpha: 1 - (Do / De), Do, De };
    }

    function pearsonCorrelation(x, y) {
        const n = x.length;
        const meanX = x.reduce((a, b) => a + b, 0) / n;
        const meanY = y.reduce((a, b) => a + b, 0) / n;
        let num = 0, dx = 0, dy = 0;
        for (let i = 0; i < n; i++) {
            const xv = x[i] - meanX;
            const yv = y[i] - meanY;
            num += xv * yv;
            dx += xv * xv;
            dy += yv * yv;
        }
        return num / Math.sqrt(dx * dy);
    }

    async function runValidationChecks() {
        const results = [];
        const total = appState.data.length;

        for (const task of appState.analysisTasks) {
            if (task.reliabilityCheck && task.reliabilityColumn) {
                const aiValsRaw = appState.data.map(r => r[task.outputColumn]);
                const humanValsRaw = appState.data.map(r => r[task.reliabilityColumn]);
                const aiSets = aiValsRaw.map(v => parseCodeSet(v));
                const humanSets = humanValsRaw.map(v => parseCodeSet(v));
                const aiVals = aiSets.map(s => s.slice().sort().join('|'));
                const humanVals = humanSets.map(s => s.slice().sort().join('|'));
                const exact = aiVals.filter((v,i) => v === humanVals[i]).length / total;
                const kappa = computeKappa(aiVals, humanVals);
                const { alpha, Do, De } = computeKrippendorffAlpha(aiSets, humanSets, task.reliabilityMethod);
                const entry = { validation: 'Intercoder Reliability', task: task.outputColumn, method: "Cohen's Kappa", score: Number(kappa.toFixed(2)), krippendorff_alpha: Number(alpha.toFixed(2)), exact_match: Number(exact.toFixed(2)), reference_column: task.reliabilityColumn, distance: task.reliabilityMethod || 'jaccard', observed_disagreement: Number(Do.toFixed(2)), expected_disagreement: Number(De.toFixed(2)) };
                results.push(entry);
                log(JSON.stringify(entry), 'VALIDATION');
            }
        }

        if (appState.validationSettings.consistencyCheck && appState.validationSettings.consistencyColumns.length >= 2) {
            const cols = appState.validationSettings.consistencyColumns;
            const c1 = appState.data.map(r => r[cols[0]]);
            const c2 = appState.data.map(r => r[cols[1]]);
            const nums1 = c1.map(Number);
            const nums2 = c2.map(Number);
            let method, corr;
            if (nums1.every(n => !isNaN(n)) && nums2.every(n => !isNaN(n))) {
                method = 'Pearson correlation';
                corr = pearsonCorrelation(nums1, nums2);
            } else {
                method = 'Pairwise agreement';
                corr = c1.filter((v,i)=>v===c2[i]).length / total;
            }
            const entry = { validation: 'Consistency Check', columns: cols, method, correlation: Number(corr.toFixed(2)) };
            results.push(entry);
            log(JSON.stringify(entry), 'VALIDATION');
        }

        if (appState.validationSettings.missingCheck) {
            for (const task of appState.analysisTasks) {
                const vals = appState.data.map(r => r[task.outputColumn]);
                const empty = vals.filter(v => v === '' || v === null || v === undefined || v === 'ERROR').length;
                const entry = { validation: 'Missing Output Check', column: task.outputColumn, empty_rows: empty, total_rows: total, percentage: Number((empty * 100 / total).toFixed(2)) };
                results.push(entry);
                log(JSON.stringify(entry), 'VALIDATION');
            }
        }

        if (appState.validationSettings.faceValidity) {
            const task = appState.analysisTasks.find(t => t.reliabilityCheck && t.reliabilityColumn);
            if (task) {
                const sampleSize = Math.min(10, total);
                const rows = [];
                for (let i = 0; i < sampleSize; i++) {
                    const idx = Math.floor(Math.random() * total);
                    const row = appState.data[idx];
                    rows.push({row_id: idx + 1, human_label: row[task.reliabilityColumn], ai_label: row[task.outputColumn]});
                }
                const csv = 'row_id,human_label,ai_label\n' + rows.map(r => `${r.row_id},"${r.human_label}","${r.ai_label}"`).join('\n');
                downloadFile(csv, 'face_validity_sample.csv', 'text/csv;charset=utf-8;');
                const entry = { validation: 'Face Validity Sample', rows: sampleSize, columns: [task.reliabilityColumn, task.outputColumn] };
                results.push(entry);
                log(JSON.stringify(entry), 'VALIDATION');
            }
        }

        if (appState.validationSettings.promptStability) {
            const sampleSize = Math.min(5, total);
            const indices = Array.from({length: sampleSize}, () => Math.floor(Math.random()*total));
            let diffTotal = 0;
            for (const idx of indices) {
                const row = appState.data[idx];
                for (const task of appState.analysisTasks) {
                    const original = row[task.outputColumn] || '';
                    const prompt = 'Please ' + buildPrompt(task, row);
                    try {
                        const result = await processWithApi(prompt, task.maxTokens);
                        diffTotal += Math.abs(estimateTokens(result.text) - estimateTokens(original));
                    } catch (e) {
                        log(`Stability check error row ${idx+1}: ${e.message}`, 'ERROR');
                    }
                }
            }
            const entry = { validation: 'Prompt Stability', rows_tested: sampleSize, avg_output_diff_tokens: Number((diffTotal / (sampleSize * appState.analysisTasks.length)).toFixed(2)), variation_method: 'syntactic filler' };
            results.push(entry);
            log(JSON.stringify(entry), 'VALIDATION');
        }

        appState.runLog.validations = results;
        renderReliabilityDashboard();
    }

    async function processWithApi(prompt, maxTokens) {
        const systemPrompt = ui.systemPromptInput.value;
        const provider = ui.providerSelector.value;
        try {
            switch (provider) {
                case 'gemini': return await processWithGemini(prompt, systemPrompt, maxTokens);
                case 'openai': return await processWithOpenAI(prompt, systemPrompt, maxTokens);
                case 'ollama': return await processWithOllama(prompt, systemPrompt, maxTokens);
                default: throw new Error(`Unknown provider: ${provider}`);
            }
        } catch (error) {
            throw error;
        }
    }
    async function processWithGemini(userPrompt, systemPrompt, maxTokens) {
        const model = ui.geminiModelSelector.value;
        const apiKey = ui.geminiApiKeyInput.value;
        const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
        const projectDesc = ui.projectDescriptionInput.value;
        const additional = ui.additionalContextInput.value;
        const finalPrompt = `${systemPrompt}\n\nProject Description:\n${projectDesc}\n\nAdditional Context:\n${additional}\n\n---\n\n${userPrompt}`;
        const params = appState.aiParams;
        const payload = {
            contents: [{ role: "user", parts: [{ text: finalPrompt }] }],
            generationConfig: {
                maxOutputTokens: parseInt(maxTokens, 10) || parseInt(params.maxTokens,10) || 150,
                temperature: parseFloat(params.temperature),
                topP: parseFloat(params.topP)
            }
        };
        const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        const result = await response.json();
        if (!response.ok || result.error) throw new Error(result.error ? result.error.message : `HTTP error! status: ${response.status}`);
        if (result.candidates?.[0]?.content.parts[0].text) {
            const text = result.candidates[0].content.parts[0].text;
            recordAuditEntry({
                timestamp: new Date().toISOString(),
                provider: 'gemini',
                model,
                prompt: finalPrompt,
                response: text,
                inputTokens: estimateTokens(finalPrompt),
                outputTokens: estimateTokens(text)
            });
            return { text };
        }
        throw new Error("No content in Gemini response.");
    }
    async function processWithOpenAI(userPrompt, systemPrompt, maxTokens) {
        const model = ui.openaiModelSelector.value;
        const apiKey = ui.openaiApiKeyInput.value;
        const projectDesc = ui.projectDescriptionInput.value;
        const additional = ui.additionalContextInput.value;
        if (!apiKey) throw new Error("OpenAI API Key is required.");
        if (!model) throw new Error("Please fetch and select an OpenAI model.");
        const params = appState.aiParams;
        const finalPrompt = `${systemPrompt}\n\nProject Description:\n${projectDesc}\n\nAdditional Context:\n${additional}\n\n${userPrompt}`;

        if (model.includes('gpt-5')) {
            const apiUrl = 'https://api.openai.com/v1/responses';
            const payload = {
                model,
                input: finalPrompt,
                max_output_tokens: parseInt(maxTokens, 10) || parseInt(params.maxTokens,10) || 150,
                temperature: parseFloat(params.temperature),
                reasoning: { effort: params.reasoning },
                text: { verbosity: params.verbosity }
            };
            const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` }, body: JSON.stringify(payload) });
            const result = await response.json();
            if (!response.ok || result.error) throw new Error(result.error ? result.error.message : `HTTP error! status: ${response.status}`);
            const text = result.output_text || result.output?.[0]?.content?.[0]?.text?.value;
            if (text) {
                recordAuditEntry({
                    timestamp: new Date().toISOString(),
                    provider: 'openai',
                    model,
                    prompt: finalPrompt,
                    response: text,
                    inputTokens: estimateTokens(finalPrompt),
                    outputTokens: estimateTokens(text)
                });
                return { text };
            }
            throw new Error("No content in OpenAI response.");
        } else {
            const apiUrl = `https://api.openai.com/v1/chat/completions`;
            const payload = {
                model: model,
                messages: [
                    { role: "system", content: systemPrompt },
                    { role: "system", content: `Project Description:\n${projectDesc}` },
                    { role: "system", content: `Additional Context:\n${additional}` },
                    { role: "user", content: userPrompt }
                ],
                max_tokens: parseInt(maxTokens, 10) || parseInt(params.maxTokens,10) || 150,
                temperature: parseFloat(params.temperature),
                top_p: parseFloat(params.topP),
                n: parseInt(params.n,10),
                frequency_penalty: parseFloat(params.freqPenalty),
                presence_penalty: parseFloat(params.presPenalty)
            };
            if(params.stop === 1) payload.stop = ["\n"];
            const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` }, body: JSON.stringify(payload) });
            const result = await response.json();
            if (!response.ok || result.error) throw new Error(result.error ? result.error.message : `HTTP error! status: ${response.status}`);
            if (result.choices?.[0]?.message.content) {
                const text = result.choices[0].message.content;
                recordAuditEntry({
                    timestamp: new Date().toISOString(),
                    provider: 'openai',
                    model,
                    prompt: finalPrompt,
                    response: text,
                    inputTokens: estimateTokens(finalPrompt),
                    outputTokens: estimateTokens(text)
                });
                return { text };
            }
            throw new Error("No content in OpenAI response.");
        }
    }
    async function processWithOllama(userPrompt, systemPrompt, maxTokens) {
        const model = ui.ollamaModelSelector.value;
        const url = ui.ollamaUrlInput.value;
        if (!model) throw new Error("Please fetch and select an Ollama model.");
        const apiUrl = new URL('/api/generate', url).href;
        const projectDesc = ui.projectDescriptionInput.value;
        const additional = ui.additionalContextInput.value;
        const params = appState.aiParams;
        const payload = {
            model: model,
            prompt: `Project Description:\n${projectDesc}\n\nAdditional Context:\n${additional}\n\n${userPrompt}`,
            system: systemPrompt,
            stream: false,
            options: {
                num_predict: parseInt(maxTokens, 10) || 150,
                temperature: parseFloat(params.temperature)
            }
        };
        if (params.ollamaReasoning) payload.options.reasoning = true;
        const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        const result = await response.json();
        if (!response.ok || result.error) throw new Error(result.error || `HTTP error! status: ${response.status}`);
        if (result.response) {
            const text = result.response;
            const finalPrompt = `${systemPrompt}\n\nProject Description:\n${projectDesc}\n\nAdditional Context:\n${additional}\n\n${userPrompt}`;
            recordAuditEntry({
                timestamp: new Date().toISOString(),
                provider: 'ollama',
                model,
                prompt: finalPrompt,
                response: text,
                inputTokens: estimateTokens(finalPrompt),
                outputTokens: estimateTokens(text)
            });
            return { text };
        }
        throw new Error("No content in Ollama response.");
    }

    function addAnalysisTask(type, taskData = {}) {
        const taskIndex = appState.analysisTasks.length;
        const taskId = taskData.id || crypto.randomUUID();
        let newTask, template;

        if (type === 'analyze') {
            newTask = Object.assign({ id: taskId, type: 'analyze', sourceColumn: '', outputColumn: `analysis_${taskIndex + 1}`, prompt: '', maxTokens: 150, reliabilityCheck: false, reliabilityColumn: '', reliabilityMethod: 'jaccard' }, taskData);
            template = ui.analyzeTaskTemplate;
        } else if (type === 'compare') {
            newTask = Object.assign({ id: taskId, type: 'compare', sourceColumns: [], outputColumn: `comparison_${taskIndex + 1}`, prompt: '', maxTokens: 150, reliabilityCheck: false, reliabilityColumn: '', reliabilityMethod: 'jaccard' }, taskData);
            template = ui.compareTaskTemplate;
        } else if (type === 'custom') {
            newTask = Object.assign({ id: taskId, type: 'custom', outputColumn: `custom_${taskIndex + 1}`, prompt: '', maxTokens: 150, reliabilityCheck: false, reliabilityColumn: '', reliabilityMethod: 'jaccard' }, taskData);
            template = ui.customTaskTemplate;
        } else if (type === "auto") {
            newTask = Object.assign({ id: taskId, type: "auto", outputColumn: `auto_${taskIndex + 1}`, prompt: '', maxTokens: 150, reliabilityCheck: false, reliabilityColumn: '', reliabilityMethod: 'jaccard' }, taskData);
            template = ui.autoTaskTemplate;
        } else return;

        appState.analysisTasks.push(newTask);

        const taskFragment = template.content.cloneNode(true);
        const taskCard = taskFragment.querySelector('.task-card');
        taskCard.dataset.taskId = taskId;

        taskCard.querySelector('input[data-type="outputColumn"]').value = newTask.outputColumn;
        taskCard.querySelector('textarea[data-type="prompt"]').value = newTask.prompt;
        taskCard.querySelector('input[data-type="maxTokens"]').value = newTask.maxTokens;
        const relChk = taskCard.querySelector('input[data-type="reliabilityCheck"]');
        const relSel = taskCard.querySelector('select[data-type="reliabilityColumn"]');
        const relMethod = taskCard.querySelector('select[data-type="reliabilityMethod"]');
        if(relChk) relChk.checked = newTask.reliabilityCheck;
        if(relSel) {
            relSel.value = newTask.reliabilityColumn;
            relSel.dataset.savedValue = newTask.reliabilityColumn;
            relSel.disabled = !newTask.reliabilityCheck;
        }
        if(relMethod) {
            relMethod.value = newTask.reliabilityMethod || 'jaccard';
            relMethod.disabled = !newTask.reliabilityCheck;
        }

        ui.taskContainer.appendChild(taskFragment);
        
        if (type === 'analyze') {
            const srcSel = taskCard.querySelector('select[data-type="sourceColumn"]');
            srcSel.value = newTask.sourceColumn;
            srcSel.dataset.savedValue = newTask.sourceColumn;
        } else if (type === 'compare') {
            if (newTask.sourceColumns.length === 0) {
                 addColumnToCompareTask(taskCard);
            } else {
                newTask.sourceColumns.forEach(sc => addColumnToCompareTask(taskCard, sc));
            }
        }
        
        updateTaskColumnDropdowns();
        if(!taskData) log(`Added new '${type}' task.`, 'INFO');
        updateCostEstimate();
    }

    function removeAnalysisTask(taskId) {
        appState.analysisTasks = appState.analysisTasks.filter(task => task.id !== taskId);
        ui.taskContainer.querySelector(`.task-card[data-task-id="${taskId}"]`)?.remove();
        log('Removed analysis task.', 'INFO');
        updateCostEstimate();
    }

    function addColumnToCompareTask(taskCard, value = '') {
        const taskId = taskCard.dataset.taskId;
        const task = appState.analysisTasks.find(t => t.id === taskId);
        if (!task) return;

        const newIndex = task.sourceColumns.length;
        if(value === '') task.sourceColumns.push(''); 

        const container = taskCard.querySelector('.source-columns-container');
        const colDiv = document.createElement('div');
        colDiv.className = 'flex items-center gap-2';
        
        const select = document.createElement('select');
        select.className = 'task-input w-full bg-gray-700 border border-gray-600 rounded-md p-2 focus:ring-2 focus:ring-blue-500 focus:outline-none';
        select.dataset.type = 'sourceColumn';
        select.dataset.index = newIndex;
        select.value = value;
        select.dataset.savedValue = value;
        colDiv.appendChild(select);

        const removeBtn = document.createElement('button');
        removeBtn.className = 'remove-source-column-btn text-gray-500 hover:text-red-400 text-2xl leading-none';
        removeBtn.innerHTML = '&times;';
        colDiv.appendChild(removeBtn);

        container.appendChild(colDiv);
        updateTaskColumnDropdowns();
    }
    
    function removeColumnFromCompareTask(taskCard, colDiv, index) {
        const taskId = taskCard.dataset.taskId;
        const task = appState.analysisTasks.find(t => t.id === taskId);
        if (task) {
            task.sourceColumns.splice(index, 1);
            colDiv.remove();
            taskCard.querySelectorAll('.source-columns-container select').forEach((sel, i) => sel.dataset.index = i);
            updateCostEstimate();
        }
    }

    function updateTaskState(taskId, field, value, index = null) {
        const task = appState.analysisTasks.find(t => t.id === taskId);
        if (!task) return;
        if (field === 'sourceColumn' && task.type === 'compare') {
            task.sourceColumns[index] = value;
        } else {
            task[field] = value;
        }
        updateCostEstimate();
    }

    function updateTaskColumnDropdowns() {
        document.querySelectorAll('.task-input[data-type="sourceColumn"]').forEach(dropdown => {
            const saved = dropdown.dataset.savedValue || '';
            const currentVal = saved || dropdown.value;
            dropdown.innerHTML = '';
            if (appState.headers.length === 0) {
                dropdown.add(new Option('Upload a file first', ''));
                dropdown.disabled = true;
                dropdown.value = '';
                dropdown.dataset.savedValue = saved;
            } else {
                dropdown.add(new Option('Select a column...', ''));
                appState.headers.forEach(header => dropdown.add(new Option(header, header)));
                dropdown.disabled = false;
                if (appState.headers.includes(currentVal)) {
                    dropdown.value = currentVal;
                    dropdown.dataset.savedValue = currentVal;
                } else {
                    dropdown.value = '';
                    dropdown.dataset.savedValue = saved;
                }
            }
        });

        document.querySelectorAll('.task-input[data-type="reliabilityColumn"]').forEach(dropdown => {
            const saved = dropdown.dataset.savedValue || '';
            const currentVal = saved || dropdown.value;
            dropdown.innerHTML = '';
            if (appState.headers.length === 0) {
                dropdown.add(new Option('Upload a file first', ''));
                dropdown.disabled = true;
                dropdown.value = '';
                dropdown.dataset.savedValue = saved;
            } else {
                dropdown.add(new Option('Select a column...', ''));
                appState.headers.forEach(h => dropdown.add(new Option(h, h)));
                dropdown.disabled = !dropdown.closest('.task-card').querySelector('input[data-type="reliabilityCheck"]').checked;
                if (appState.headers.includes(currentVal)) {
                    dropdown.value = currentVal;
                    dropdown.dataset.savedValue = currentVal;
                } else {
                    dropdown.value = '';
                    dropdown.dataset.savedValue = saved;
                }
            }
        });

        document.querySelectorAll('.task-input[data-type="reliabilityMethod"]').forEach(dropdown => {
            dropdown.disabled = !dropdown.closest('.task-card').querySelector('input[data-type="reliabilityCheck"]').checked;
            const current = dropdown.value || 'jaccard';
            dropdown.innerHTML = '';
            dropdown.add(new Option('Alpha Weighted (Jaccard)', 'jaccard'));
            dropdown.add(new Option('Basic Alpha', 'binary'));
            dropdown.add(new Option('Alpha Weighted (MASI)', 'masi'));
            dropdown.value = current;
        });

        if (ui.consistencyColumns) {
            const current = Array.from(ui.consistencyColumns.selectedOptions).map(o => o.value);
            ui.consistencyColumns.innerHTML = '';
            appState.headers.forEach(h => ui.consistencyColumns.add(new Option(h, h)));
            current.forEach(val => { if (appState.headers.includes(val)) ui.consistencyColumns.querySelector(`option[value="${val}"]`).selected = true; });
        }

        updatePresetDropdowns();
    }

    function populateModelSelector(provider, models) {
        let selector = provider === 'openai' ? ui.openaiModelSelector : ui.ollamaModelSelector;
        selector.innerHTML = '';
        if (models.length === 0) { selector.add(new Option('No models found', '')); selector.disabled = true; return; }
        models.forEach(model => {
            const modelId = provider === 'openai' ? model.id : model.name;
            selector.add(new Option(modelId, modelId));
        });
        if (provider === 'openai' && models.some(m => m.id === 'gpt-4o-mini')) selector.value = 'gpt-4o-mini';
        else if (provider === 'ollama' && models.some(m => m.name.includes('llama3'))) selector.value = models.find(m => m.name.includes('llama3')).name;
        selector.disabled = false;
        updateCostEstimate();
        updateProviderSpecificParams();
    }
    
    async function autoGenerateTask() {
        if (appState.data.length === 0) { log('Upload a file before using Auto-Generate.', 'ERROR'); return; }
        if (!confirm('This will use API resources to auto-generate tasks for empty columns. Continue?')) return;
        log('Generating auto tasks...', 'API');

        const headers = appState.headers;
        const emptyCols = headers.filter(h => appState.data.every(r => !r[h]));
        if (emptyCols.length === 0) { log('No empty columns detected.', 'ERROR'); return; }
        const dataCols = headers.filter(h => !emptyCols.includes(h));
        const sampleRows = appState.data.slice(0,5);
        const preview = [headers.join(' | ')].concat(sampleRows.map(r => headers.map(h => r[h]).join(' | '))).join('\n');
        const projectDesc = ui.projectDescriptionInput.value || 'N/A';
        const additional = ui.additionalContextInput.value || '';
        const guidance = `You are configuring analysis tasks. For each task field:\n- Use {{ColumnName}} to reference the current row's data for a given column.\n- Specify exactly which column to write results into.\n- Output only the analysis text—no greetings or extra commentary, don't duplicate data from other fields.`;
        const userPrompt = `${guidance}\n\nProject Description: ${projectDesc}\nAdditional Context:\n${additional}\nSpreadsheet Preview:\n${preview}\n\nExisting columns with data: ${dataCols.join(', ')}.\nColumns needing analysis: ${emptyCols.join(', ')}.\n\nFor each column needing analysis, provide one line in the form "Column -> instruction" describing how to analyze each row.`;

        try {
            const result = await processWithApi(userPrompt, 400);
            const lines = result.text.split(/\n+/).map(l => l.trim()).filter(Boolean);
            let added = 0;
            lines.forEach(line => {
                const parts = line.split('->');
                if (parts.length >= 2) {
                    const col = parts[0].trim();
                    const instr = parts.slice(1).join('->').trim();
                    if (emptyCols.includes(col)) {
                        addAnalysisTask('custom', { outputColumn: col, prompt: instr });
                        added++;
                    }
                }
            });
            if (added > 0) log(`Auto-generated ${added} task(s).`, 'SUCCESS');
            else log('Auto generation produced no usable tasks.', 'ERROR');
        } catch (error) {
            log(`Failed to auto-generate tasks: ${error.message}`, 'ERROR');
        }
    }

    function openPresetModal(key) {
        const preset = TASK_PRESETS[key];
        if (!preset) return;
        appState.currentPreset = key;
        ui.presetModalTitle.textContent = `Configure: ${preset.label}`;
        ui.presetModalBody.innerHTML = '';
        if (preset.type === 'analyze') {
            const label = document.createElement('label');
            label.className = 'block text-sm font-medium text-gray-400 mb-1';
            label.textContent = 'Select Source Column';
            const select = document.createElement('select');
            select.id = 'preset-source-1';
            select.className = 'w-full bg-gray-700 border border-gray-600 rounded-md p-2';
            ui.presetModalBody.appendChild(label);
            ui.presetModalBody.appendChild(select);
        } else if (preset.type === 'compare') {
            ['First Column', 'Second Column'].forEach((txt, i) => {
                const label = document.createElement('label');
                label.className = 'block text-sm font-medium text-gray-400 mt-2 mb-1';
                label.textContent = txt;
                const select = document.createElement('select');
                select.id = `preset-source-${i+1}`;
                select.className = 'w-full bg-gray-700 border border-gray-600 rounded-md p-2';
                ui.presetModalBody.appendChild(label);
                ui.presetModalBody.appendChild(select);
            });
        } else {
            const p = document.createElement('p');
            p.className = 'text-sm text-gray-300';
            p.textContent = 'This task has no required source columns.';
            ui.presetModalBody.appendChild(p);
        }
        updatePresetDropdowns();
        ui.presetModal.classList.remove('hidden');
    }

    function closePresetModal() {
        appState.currentPreset = null;
        ui.presetModal.classList.add('hidden');
    }

    function updatePresetDropdowns() {
        ui.presetModalBody.querySelectorAll('select').forEach(sel => {
            const currentVal = sel.value;
            sel.innerHTML = '';
            if (appState.headers.length === 0) {
                sel.add(new Option('Upload a file first', ''));
                sel.disabled = true;
            } else {
                sel.add(new Option('Select column...', ''));
                appState.headers.forEach(h => sel.add(new Option(h, h)));
                sel.disabled = false;
                if (appState.headers.includes(currentVal)) sel.value = currentVal;
            }
        });
    }

    function confirmPresetTask() {
        const key = appState.currentPreset;
        const preset = TASK_PRESETS[key];
        if (!preset) return;
        let taskData = { outputColumn: preset.outputColumn, maxTokens: 150, presetLabel: preset.label };
        let prompt = preset.promptTemplate;
        if (preset.type === 'analyze') {
            const col = document.getElementById('preset-source-1').value;
            if (!col) { alert('Select a column'); return; }
            taskData.sourceColumn = col;
            prompt = prompt.replace('{{COLUMN}}', `{{${col}}}`);
        } else if (preset.type === 'compare') {
            const c1 = document.getElementById('preset-source-1').value;
            const c2 = document.getElementById('preset-source-2').value;
            if (!c1 || !c2) { alert('Select two columns'); return; }
            taskData.sourceColumns = [c1, c2];
            prompt = prompt.replace('{{COLUMNS_DATA}}', `{{${c1}}}\n{{${c2}}}`);
        }
        taskData.prompt = prompt;
        addAnalysisTask(preset.type, taskData);
        closePresetModal();
        const cards = ui.taskContainer.querySelectorAll('.task-card');
        if (cards.length) cards[cards.length-1].scrollIntoView({behavior:'smooth'});
    }

    function gatherGuessContext() {
        const headers = appState.headers;
        const sampleRows = appState.data.slice(0,3);
        const preview = headers.length ? [headers.join(' | ')].concat(sampleRows.map(r => headers.map(h => r[h]).join(' | '))).join('\n') : '';
        const tasks = appState.analysisTasks.map(t => `${t.outputColumn}: ${t.prompt}`).join('\n');
        return {
            system: ui.systemPromptInput.value,
            project: ui.projectDescriptionInput.value,
            additional: ui.additionalContextInput.value,
            preview,
            tasks
        };
    }

    function toggleGuessLoading(btn, loading) {
        if (loading) {
            btn.disabled = true;
            btn.dataset.label = btn.textContent;
            btn.innerHTML = '<svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24" stroke="currentColor"><circle cx="12" cy="12" r="10" stroke-width="4" class="opacity-25"/><path stroke-linecap="round" stroke-width="4" d="M4 12a8 8 0 018-8" class="opacity-75"/></svg>';
        } else {
            btn.disabled = false;
            btn.innerHTML = btn.dataset.label || 'Guess';
        }
    }

    async function guessProjectOverview() {
        const btn = ui.guessProjectBtn;
        const errorEl = btn.parentElement.querySelector('.guess-error');
        errorEl.classList.add('hidden');
        toggleGuessLoading(btn, true);
        const ctx = gatherGuessContext();
        const prompt = `${ctx.system}\n\nSpreadsheet Preview:\n${ctx.preview}\n\nExisting Tasks:\n${ctx.tasks}\n\nProvide a short project overview describing what this spreadsheet is about.`;
        try {
            const res = await processWithApi(prompt, 150);
            ui.projectDescriptionInput.value = res.text.trim();
            updateCostEstimate();
        } catch (err) {
            errorEl.classList.remove('hidden');
        } finally {
            toggleGuessLoading(btn, false);
        }
    }

    async function guessAdditionalContext() {
        const btn = ui.guessAdditionalBtn;
        const errorEl = btn.parentElement.querySelector('.guess-error');
        errorEl.classList.add('hidden');
        toggleGuessLoading(btn, true);
        const ctx = gatherGuessContext();
        const prompt = `${ctx.system}\n\nProject Overview:\n${ctx.project}\nSpreadsheet Preview:\n${ctx.preview}\nExisting Tasks:\n${ctx.tasks}\n\nSuggest brief additional context to improve the analysis.`;
        try {
            const res = await processWithApi(prompt, 150);
            ui.additionalContextInput.value = res.text.trim();
            updateCostEstimate();
        } catch (err) {
            errorEl.classList.remove('hidden');
        } finally {
            toggleGuessLoading(btn, false);
        }
    }

    async function guessTaskPrompt(taskCard, btn) {
        const errorEl = btn.parentElement.querySelector('.guess-error');
        errorEl.classList.add('hidden');
        toggleGuessLoading(btn, true);
        const taskId = taskCard.dataset.taskId;
        const task = appState.analysisTasks.find(t => t.id === taskId);
        const ctx = gatherGuessContext();
        const base = 'You are configuring analysis tasks. For each task field:\n- Use {{ColumnName}} to reference the current row\'s data for a given column.\n- Specify exactly which column to write results into.\n- Output only the analysis text—no greetings or extra commentary, don\'t duplicate data from other fields.';
        let desc = '';
        if (task.type === 'analyze') desc = `Analyze column ${task.sourceColumn} and save to ${task.outputColumn}.`;
        else if (task.type === 'compare') desc = `Compare columns ${task.sourceColumns.join(', ')} and save to ${task.outputColumn}.`;
        else desc = `Write results to column ${task.outputColumn}.`;
        const prompt = `${base}\n\nProject Overview:\n${ctx.project}\nAdditional Context:\n${ctx.additional}\nSpreadsheet Preview:\n${ctx.preview}\nExisting Tasks:\n${ctx.tasks}\n\n${desc}\nProvide only the instructions.`;
        try {
            const res = await processWithApi(prompt, 200);
            taskCard.querySelector('textarea[data-type="prompt"]').value = res.text.trim();
            updateTaskState(taskId, 'prompt', res.text.trim());
        } catch (err) {
            errorEl.classList.remove('hidden');
        } finally {
            toggleGuessLoading(btn, false);
        }
    }
    function startTestMode(task) {
        ui.selectTaskModal.classList.add('hidden');
        appState.testMode = { isActive: true, task, currentIndex: 0 };
        ui.testModeTitle.textContent = `Test Task -> ${task.outputColumn}`;
        renderTestModeView();
        ui.testModeModal.classList.remove('hidden');
    }

    function renderTestModeView() {
        const { task, currentIndex } = appState.testMode;
        if (!task || !appState.data[currentIndex]) return;
        const row = appState.data[currentIndex];
        const prompt = buildPrompt(task, row);
        
        let rowDataText;
        if (task.type === 'analyze') {
            rowDataText = `Column '${task.sourceColumn}':\n\n${row[task.sourceColumn] || ''}`;
        } else if (task.type === 'compare') {
            rowDataText = task.sourceColumns.map(sc => `Column '${sc}': ${row[sc] || ''}`).join('\n');
        } else { // Custom task
            rowDataText = 'N/A (Custom Task references columns in prompt)';
        }

        ui.testModeNavInfo.textContent = `Row ${currentIndex + 1} of ${appState.data.length}`;
        ui.testModeRowData.textContent = rowDataText;
        ui.testModePrompt.textContent = prompt;
        ui.testModeOutput.textContent = 'Click "Rerun" to generate output.';

        ui.testModePrevBtn.disabled = currentIndex === 0;
        ui.testModeNextBtn.disabled = currentIndex >= appState.data.length - 1;
    }

    async function runTestModeRow() {
        const { task, currentIndex } = appState.testMode;
        if (!task || !appState.data[currentIndex]) return;

        const row = appState.data[currentIndex];
        const prompt = buildPrompt(task, row);

        ui.testModeOutput.textContent = 'Processing...';
        try {
            const result = await processWithApi(prompt, task.maxTokens);
            ui.testModeOutput.textContent = result.text;
        } catch (error) {
            ui.testModeOutput.textContent = `ERROR: ${error.message}`;
        }
    }

    function saveProfile() {
        log('Saving profile...', 'PROFILE');
        const profile = {
            provider: ui.providerSelector.value,
            modelConfig: {},
            systemPrompt: ui.systemPromptInput.value,
            projectDescription: ui.projectDescriptionInput.value,
            additionalContext: ui.additionalContextInput.value,
            analysisTasks: appState.analysisTasks,
            includeRowContext: appState.includeRowContext,
            validationSettings: appState.validationSettings,
            aiParams: appState.aiParams,
        };

        switch (profile.provider) {
            case 'openai':
                profile.modelConfig.model = ui.openaiModelSelector.value;
                profile.modelConfig.apiKey = ui.openaiApiKeyInput.value;
                if(profile.modelConfig.apiKey) log('WARNING: Saving API key to a file is a security risk. Handle with care.', 'ERROR');
                break;
            case 'ollama':
                profile.modelConfig.model = ui.ollamaModelSelector.value;
                profile.modelConfig.url = ui.ollamaUrlInput.value;
                break;
             case 'gemini':
                profile.modelConfig.model = ui.geminiModelSelector.value;
                profile.modelConfig.apiKey = ui.geminiApiKeyInput.value;
                break;
        }
        
        downloadFile(JSON.stringify(profile, null, 2), 'ai-processor-profile.json', 'application/json');
    }

    async function loadProfile(event) {
        const file = event.target.files[0];
        if (!file) return;

        log(`Loading profile: ${file.name}`, 'PROFILE');
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const profile = JSON.parse(e.target.result);
                
                // --- Reset current state ---
                ui.taskContainer.innerHTML = '';
                appState.analysisTasks = [];

                // --- Apply model config ---
                ui.providerSelector.value = profile.provider || 'gemini';
                ui.providerSelector.dispatchEvent(new Event('change'));

                ui.systemPromptInput.value = profile.systemPrompt || DEFAULT_SYSTEM_PROMPT;
                ui.projectDescriptionInput.value = profile.projectDescription || '';
                ui.additionalContextInput.value = profile.additionalContext || '';

                const { modelConfig } = profile;
                if (modelConfig) {
                    switch (profile.provider) {
                        case 'openai':
                            ui.openaiApiKeyInput.value = modelConfig.apiKey || '';
                            if(modelConfig.apiKey && modelConfig.model) {
                                await ui.fetchOpenAIModelsBtn.click();
                                setTimeout(() => { ui.openaiModelSelector.value = modelConfig.model; updateCostEstimate(); updateProviderSpecificParams(); }, 500);
                            }
                            break;
                        case 'ollama':
                            ui.ollamaUrlInput.value = modelConfig.url || 'http://localhost:11434';
                            if (modelConfig.model) {
                                await ui.fetchOllamaModelsBtn.click();
                                setTimeout(() => { ui.ollamaModelSelector.value = modelConfig.model; updateCostEstimate(); }, 500);
                            }
                            break;
                         case 'gemini':
                            ui.geminiModelSelector.value = modelConfig.model || 'gemini-1.5-flash-latest';
                            ui.geminiApiKeyInput.value = modelConfig.apiKey || '';
                            break;
                    }
                }

                appState.includeRowContext = profile.includeRowContext !== false;
                ui.includeRowContextToggle.checked = appState.includeRowContext;
                if (profile.validationSettings) {
                    appState.validationSettings = Object.assign(appState.validationSettings, profile.validationSettings);
                    ui.checkMissingOutputs.checked = appState.validationSettings.missingCheck;
                    ui.consistencyCheck.checked = appState.validationSettings.consistencyCheck;
                    ui.promptStability.checked = appState.validationSettings.promptStability;
                    ui.faceValidity.checked = appState.validationSettings.faceValidity;
                    updateTaskColumnDropdowns();
                }

                if (profile.aiParams) {
                    appState.aiParams = Object.assign(appState.aiParams, profile.aiParams);
                    ui.cachedInputsToggle.checked = appState.aiParams.cache;
                    ui.temperatureSlider.value = appState.aiParams.temperature;
                    ui.temperatureValue.textContent = appState.aiParams.temperature;
                    ui.maxTokensSlider.value = appState.aiParams.maxTokens;
                    ui.maxTokensValue.textContent = appState.aiParams.maxTokens;
                    ui.topPSlider.value = appState.aiParams.topP;
                    ui.topPValue.textContent = appState.aiParams.topP;
                    ui.nSlider.value = appState.aiParams.n;
                    ui.nValue.textContent = appState.aiParams.n;
                    ui.stopSlider.value = appState.aiParams.stop;
                    ui.stopValue.textContent = appState.aiParams.stop;
                    ui.freqPenaltySlider.value = appState.aiParams.freqPenalty;
                    ui.freqPenaltyValue.textContent = appState.aiParams.freqPenalty;
                    ui.presPenaltySlider.value = appState.aiParams.presPenalty;
                    ui.presPenaltyValue.textContent = appState.aiParams.presPenalty;
                    ui.gpt5ReasoningSelect.value = appState.aiParams.reasoning || 'minimal';
                    ui.gpt5VerbositySelect.value = appState.aiParams.verbosity || 'medium';
                    ui.ollamaReasoningToggle.checked = appState.aiParams.ollamaReasoning || false;
                }

                // --- Apply analysis tasks ---
                if (profile.analysisTasks && Array.isArray(profile.analysisTasks)) {
                    profile.analysisTasks.forEach(task => addAnalysisTask(task.type, task));
                }

                updateCostEstimate();
                updateProviderSpecificParams();
                log('Profile loaded successfully.', 'SUCCESS');

            } catch (error) {
                log(`Failed to load profile. Invalid JSON. ${error.message}`, 'ERROR');
            }
        };
        reader.readAsText(file);
        event.target.value = '';
    }

    // --- Event Listeners ---
    ui.fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;
        appState.file = file;
        ui.fileInfo.textContent = `Selected: ${file.name}`;
        log(`File selected: ${file.name}`, 'FILE');
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                appState.data = results.data;
                appState.headers = results.meta.fields;
                renderDataPreview();
                renderAnalysisDashboard();
                updateTaskColumnDropdowns();
                updateCostEstimate();
                log(`Parsed ${appState.data.length} rows with columns: ${appState.headers.join(', ')}`, 'FILE');
            },
            error: (err) => { log(`File parsing error: ${err.message}`, 'ERROR'); }
        });
    });
    
    function updateProviderSpecificParams() {
        ui.gpt5Options.classList.add('hidden');
        ui.ollamaOptions.classList.add('hidden');
        if (ui.providerSelector.value === 'openai' && ui.openaiModelSelector.value.includes('gpt-5')) {
            ui.gpt5Options.classList.remove('hidden');
        }
        if (ui.providerSelector.value === 'ollama') {
            ui.ollamaOptions.classList.remove('hidden');
        }
    }

    ui.providerSelector.addEventListener('change', () => {
        document.querySelectorAll('.config-group').forEach(el => el.classList.add('hidden-config'));
        document.getElementById(`${ui.providerSelector.value}-config`).classList.remove('hidden-config');
        updateCostEstimate();
        updateProviderSpecificParams();
    });

    ui.openaiModelSelector.addEventListener('change', updateProviderSpecificParams);

    [ui.openaiParamsBtn, ui.geminiParamsBtn, ui.ollamaParamsBtn].forEach(btn => {
        btn.addEventListener('click', () => ui.aiParamsPanel.classList.toggle('hidden'));
    });

    ui.cachedInputsToggle.addEventListener('change', e => { appState.aiParams.cache = e.target.checked; updateCostEstimate(); });
    ui.temperatureSlider.addEventListener('input', e => {
        appState.aiParams.temperature = parseFloat(e.target.value);
        ui.temperatureValue.textContent = e.target.value;
    });
    ui.maxTokensSlider.addEventListener('input', e => {
        appState.aiParams.maxTokens = parseInt(e.target.value,10);
        ui.maxTokensValue.textContent = e.target.value;
        updateCostEstimate();
    });
    ui.topPSlider.addEventListener('input', e => {
        appState.aiParams.topP = parseFloat(e.target.value);
        ui.topPValue.textContent = e.target.value;
    });
    ui.nSlider.addEventListener('input', e => {
        appState.aiParams.n = parseInt(e.target.value,10);
        ui.nValue.textContent = e.target.value;
    });
    ui.stopSlider.addEventListener('input', e => {
        appState.aiParams.stop = parseInt(e.target.value,10);
        ui.stopValue.textContent = e.target.value;
    });
    ui.freqPenaltySlider.addEventListener('input', e => {
        appState.aiParams.freqPenalty = parseFloat(e.target.value);
        ui.freqPenaltyValue.textContent = e.target.value;
    });
    ui.presPenaltySlider.addEventListener('input', e => {
        appState.aiParams.presPenalty = parseFloat(e.target.value);
        ui.presPenaltyValue.textContent = e.target.value;
    });

    ui.gpt5ReasoningSelect.addEventListener('change', e => { appState.aiParams.reasoning = e.target.value; });
    ui.gpt5VerbositySelect.addEventListener('change', e => { appState.aiParams.verbosity = e.target.value; });
    ui.ollamaReasoningToggle.addEventListener('change', e => { appState.aiParams.ollamaReasoning = e.target.checked; });
    
    [ui.geminiModelSelector, ui.openaiModelSelector, ui.ollamaModelSelector].forEach(el => {
        el.addEventListener('input', updateCostEstimate);
        el.addEventListener('change', updateCostEstimate);
    });
    [ui.systemPromptInput, ui.projectDescriptionInput, ui.additionalContextInput].forEach(el => el.addEventListener('input', updateCostEstimate));
    
    ui.fetchOpenAIModelsBtn.addEventListener('click', async () => {
        const apiKey = ui.openaiApiKeyInput.value;
        if (!apiKey) { log('Please enter an OpenAI API key first.', 'ERROR'); return; }
        log('Fetching OpenAI models...', 'API');
        ui.fetchOpenAIModelsBtn.disabled = true;
        ui.openaiModelSelector.disabled = true;
        ui.openaiModelSelector.innerHTML = '<option>Fetching...</option>';
        try {
            const response = await fetch('https://api.openai.com/v1/models', { headers: { 'Authorization': `Bearer ${apiKey}` } });
            if (!response.ok) { const err = await response.json(); throw new Error(err.error.message); }
            const result = await response.json();
            const gptModels = result.data.filter(m => m.id.includes('gpt')).sort((a, b) => a.id.localeCompare(b.id));
            appState.availableModels.openai = gptModels;
            populateModelSelector('openai', gptModels);
            log(`Successfully fetched ${gptModels.length} OpenAI models.`, 'SUCCESS');
        } catch (error) {
            log(`Failed to fetch OpenAI models: ${error.message}`, 'ERROR');
            ui.openaiModelSelector.innerHTML = '<option>Failed to fetch</option>';
        } finally {
            ui.fetchOpenAIModelsBtn.disabled = false;
        }
    });

    async function listOllamaModels(baseURL) {
        try {
            const r = await fetch(new URL('/v1/models', baseURL).href, {
                method: 'GET',
                mode: 'cors'
            });
            if (r.ok) {
                const j = await r.json();
                if (Array.isArray(j.data)) {
                    return j.data.map(m => ({ name: m.id }));
                }
            }
        } catch (err) {
            console.error('Error fetching Ollama models via /v1/models:', err);
        }

        const r2 = await fetch(new URL('/api/tags', baseURL).href, {
            method: 'GET',
            mode: 'cors'
        });
        if (!r2.ok) throw new Error(`Server returned status ${r2.status}`);
        const j2 = await r2.json();
        return j2.models ?? [];
    }

    ui.fetchOllamaModelsBtn.addEventListener('click', async () => {
        const url = ui.ollamaUrlInput.value;
        log('Fetching Ollama models...', 'API');
        ui.fetchOllamaModelsBtn.disabled = true;
        ui.ollamaModelSelector.disabled = true;
        ui.ollamaModelSelector.innerHTML = '<option>Fetching...</option>';
        try {
            const models = (await listOllamaModels(url)).sort((a, b) => a.name.localeCompare(b.name));
            appState.availableModels.ollama = models;
            populateModelSelector('ollama', models);
            log(`Successfully fetched ${models.length} local models.`, 'SUCCESS');
        } catch (error) {
            log(`Failed to fetch Ollama models. Is Ollama running at ${url}? Details: ${error.message}`, 'ERROR');
            ui.ollamaModelSelector.innerHTML = '<option>Failed to fetch</option>';
        } finally {
            ui.fetchOllamaModelsBtn.disabled = false;
            ui.ollamaModelSelector.disabled = false;
        }
    });

    ui.addTaskBtn.addEventListener('click', () => ui.addTaskMenu.classList.toggle('hidden'));

    ui.addPresetBtn.addEventListener('click', () => ui.addPresetMenu.classList.toggle('hidden'));

    document.addEventListener('click', (e) => {
        if (!ui.addTaskDropdownContainer.contains(e.target)) {
            ui.addTaskMenu.classList.add('hidden');
        }
        if (!ui.addPresetDropdownContainer.contains(e.target)) {
            ui.addPresetMenu.classList.add('hidden');
        }
    });

    ui.addTaskMenu.addEventListener('click', (e) => {
        e.preventDefault();
        const taskType = e.target.dataset.taskType;
        if (taskType) {
            ui.addTaskMenu.classList.add("hidden");
            if (taskType === "auto") autoGenerateTask();
            else addAnalysisTask(taskType);
        }
    });

    ui.addPresetMenu.addEventListener('click', (e) => {
        e.preventDefault();
        const presetKey = e.target.dataset.presetKey;
        if (presetKey) {
            ui.addPresetMenu.classList.add('hidden');
            openPresetModal(presetKey);
        }
    });
    ui.taskContainer.addEventListener('click', (e) => {
        const taskCard = e.target.closest('.task-card');
        if (!taskCard) return;
        if (e.target.classList.contains('remove-task-btn')) removeAnalysisTask(taskCard.dataset.taskId);
        else if (e.target.classList.contains('add-source-column-btn')) addColumnToCompareTask(taskCard);
        else if (e.target.classList.contains('remove-source-column-btn')) {
            const colDiv = e.target.closest('.flex');
            const select = colDiv.querySelector('select');
            removeColumnFromCompareTask(taskCard, colDiv, parseInt(select.dataset.index));
        }
    });
    
    ui.taskContainer.addEventListener('input', (e) => {
        const taskCard = e.target.closest('.task-card');
        const dataType = e.target.dataset.type;
        if(taskCard && dataType) {
            const val = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
            updateTaskState(taskCard.dataset.taskId, dataType, val, e.target.dataset.index ? parseInt(e.target.dataset.index) : null);
            if(e.target.dataset.type === 'reliabilityCheck') {
                const sel = taskCard.querySelector('select[data-type="reliabilityColumn"]');
                const methodSel = taskCard.querySelector('select[data-type="reliabilityMethod"]');
                if(sel) sel.disabled = !e.target.checked;
                if(methodSel) methodSel.disabled = !e.target.checked;
            }
            if(dataType === 'sourceColumn' || dataType === 'reliabilityColumn') {
                e.target.dataset.savedValue = val;
            }
        }
    });

    ui.taskContainer.addEventListener('change', (e) => {
        const taskCard = e.target.closest('.task-card');
        const dataType = e.target.dataset.type;
        if(taskCard && dataType) {
            const val = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
            updateTaskState(taskCard.dataset.taskId, dataType, val, e.target.dataset.index ? parseInt(e.target.dataset.index) : null);
            if(e.target.dataset.type === 'reliabilityCheck') {
                const sel = taskCard.querySelector('select[data-type="reliabilityColumn"]');
                const methodSel = taskCard.querySelector('select[data-type="reliabilityMethod"]');
                if(sel) sel.disabled = !e.target.checked;
                if(methodSel) methodSel.disabled = !e.target.checked;
            }
            if(dataType === 'sourceColumn' || dataType === 'reliabilityColumn') {
                e.target.dataset.savedValue = val;
            }
        }
    });

    ui.runBtn.addEventListener('click', async () => {
        if (appState.isProcessing || appState.data.length === 0) { log('Cannot start run. No data loaded.', 'ERROR'); return; }
        if (appState.analysisTasks.length === 0) { log('Cannot start run. Please add at least one analysis task.', 'ERROR'); return; }
        for (const task of appState.analysisTasks) {
            if (!task.outputColumn || !task.prompt ||
               (task.type === 'analyze' && !task.sourceColumn) ||
               (task.type === 'compare' && task.sourceColumns.some(sc => !sc))) {
                log(`Task is incomplete. Please fill all fields for each task before running.`, 'ERROR'); return;
            }
        }

        if (window.innerWidth < 768) {
            ui.rightPanel.classList.remove('hidden');
        }
        
        setProcessingState(true);
        appState.runLog.startTime = new Date().toISOString();
        appState.runLog.aiParams = { ...appState.aiParams };
        log(`Starting pipeline with ${appState.analysisTasks.length} tasks...`, 'RUN');
        
        let processedData = JSON.parse(JSON.stringify(appState.data));
        const newHeaders = [...appState.headers];
        let totalItems = 0;
        const totalItemsToProcess = appState.data.length * appState.analysisTasks.length;

        for (const task of appState.analysisTasks) {
            if (!newHeaders.includes(task.outputColumn)) newHeaders.push(task.outputColumn);
            appState.runLog.entries.push({ task: task.outputColumn, row_context_included: appState.includeRowContext, timestamp: new Date().toISOString() });
            
            for (let i = 0; i < processedData.length; i++) {
                const row = processedData[i];
                const prompt = buildPrompt(task, row);
                try {
                    const result = await processWithApi(prompt, task.maxTokens);
                    row[task.outputColumn] = result.text.trim();
                } catch (error) {
                    log(`ERROR on row ${i + 1}, task '${task.outputColumn}': ${error.message}`, 'ERROR');
                    row[task.outputColumn] = 'ERROR';
                }
                totalItems++;
                ui.progressBar.style.width = `${(totalItems / totalItemsToProcess) * 100}%`;
            }
        }
        
        appState.data = processedData;
        appState.headers = newHeaders;
        await runValidationChecks();
        renderDataPreview();
        renderAnalysisDashboard();
        log('Pipeline finished successfully!', 'SUCCESS');
        setProcessingState(false);
        ui.downloadResultsBtn.style.display = 'inline-block';
    });
    
    ui.downloadResultsBtn.addEventListener('click', () => {
        if (!appState.data || appState.data.length === 0) return;
        const csvContent = Papa.unparse(appState.data);
        const originalFilename = appState.file.name.replace(/\.[^/.]+$/, "");
        downloadFile(csvContent, `${originalFilename}_processed.csv`, 'text/csv;charset=utf-8;');
    });

    ui.downloadLogBtn.addEventListener('click', () => {
        if (!appState.runLog || appState.runLog.entries.length === 0) return;
        const logContent = JSON.stringify(appState.runLog, null, 2);
        const originalFilename = appState.file ? appState.file.name.replace(/\.[^/.]+$/, "") : 'log';
        downloadFile(logContent, `${originalFilename}_auditlog.json`, 'application/json');
    });

    ui.testModeBtn.addEventListener('click', () => {
        if (appState.data.length === 0) { log('Load a file to enter Test Mode.', 'ERROR'); return; }
        const validTasks = appState.analysisTasks.filter(t => t.outputColumn && ((t.type === 'analyze' && t.sourceColumn) || (t.type === 'compare' && t.sourceColumns.length > 0 && !t.sourceColumns.some(sc => !sc)) || t.type === 'custom' || t.type === 'auto'));
        if (validTasks.length === 0) { log('Please fully configure at least one analysis task.', 'ERROR'); return; }
        ui.testableTasksList.innerHTML = '';
        validTasks.forEach(task => {
            const button = document.createElement('button');
            button.className = 'w-full text-left bg-gray-700 hover:bg-gray-600 text-white font-bold py-2 px-4 rounded-md';
            button.textContent = `Test Task -> ${task.outputColumn}`;
            button.onclick = () => startTestMode(task);
            ui.testableTasksList.appendChild(button);
        });
        ui.selectTaskModal.classList.remove('hidden');
    });

    ui.closeSelectTaskModalBtn.addEventListener('click', () => ui.selectTaskModal.classList.add('hidden'));
    ui.closeTestModeBtn.addEventListener('click', () => ui.testModeModal.classList.add('hidden'));
    ui.testModeRerunBtn.addEventListener('click', runTestModeRow);
    ui.testModeNextBtn.addEventListener('click', () => { if (appState.testMode.currentIndex < appState.data.length - 1) { appState.testMode.currentIndex++; renderTestModeView(); } });
    ui.testModePrevBtn.addEventListener('click', () => { if (appState.testMode.currentIndex > 0) { appState.testMode.currentIndex--; renderTestModeView(); } });
    ui.guessProjectBtn.addEventListener('click', guessProjectOverview);
    ui.guessAdditionalBtn.addEventListener('click', guessAdditionalContext);
    ui.taskContainer.addEventListener('click', (e) => {
        if(e.target.classList.contains('guess-task-btn')) {
            const card = e.target.closest('.task-card');
            if(card) guessTaskPrompt(card, e.target);
        }
    });
    ui.confirmPresetBtn.addEventListener('click', confirmPresetTask);
    ui.cancelPresetBtn.addEventListener('click', closePresetModal);
    ui.includeRowContextToggle.addEventListener('change', e => { appState.includeRowContext = e.target.checked; updateCostEstimate(); });
    ui.checkMissingOutputs.addEventListener('change', e => appState.validationSettings.missingCheck = e.target.checked);
    ui.consistencyCheck.addEventListener('change', e => appState.validationSettings.consistencyCheck = e.target.checked);
    ui.consistencyColumns.addEventListener('change', e => { appState.validationSettings.consistencyColumns = Array.from(e.target.selectedOptions).map(o => o.value); });
    ui.faceValidity.addEventListener('change', e => appState.validationSettings.faceValidity = e.target.checked);
    ui.promptStability.addEventListener('change', e => appState.validationSettings.promptStability = e.target.checked);
    ui.saveProfileBtn.addEventListener('click', saveProfile);
    ui.loadProfileInput.addEventListener('change', loadProfile);

    // --- Initial Load ---
    addAnalysisTask('analyze');
    ui.providerSelector.dispatchEvent(new Event('change'));
    ui.temperatureValue.textContent = ui.temperatureSlider.value;
    ui.maxTokensValue.textContent = ui.maxTokensSlider.value;
    ui.topPValue.textContent = ui.topPSlider.value;
    ui.nValue.textContent = ui.nSlider.value;
    ui.stopValue.textContent = ui.stopSlider.value;
    ui.freqPenaltyValue.textContent = ui.freqPenaltySlider.value;
    ui.presPenaltyValue.textContent = ui.presPenaltySlider.value;
    ui.downloadLogBtn.style.display = 'inline-block';
    ui.downloadLogBtn.disabled = appState.runLog.entries.length === 0;
});