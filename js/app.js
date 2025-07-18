document.addEventListener('DOMContentLoaded', () => {

    // --- DOM Elements ---
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
        geminiModelSelector: document.getElementById('gemini-model-selector'),
        openaiApiKeyInput: document.getElementById('openai-api-key-input'),
        fetchOpenAIModelsBtn: document.getElementById('fetch-openai-models-btn'),
        openaiModelSelector: document.getElementById('openai-model-selector'),
        ollamaUrlInput: document.getElementById('ollama-url-input'),
        fetchOllamaModelsBtn: document.getElementById('fetch-ollama-models-btn'),
        ollamaModelSelector: document.getElementById('ollama-model-selector'),
        systemPromptInput: document.getElementById('system-prompt-input'),
        addTaskDropdownContainer: document.getElementById('add-task-dropdown-container'),
        addTaskBtn: document.getElementById('add-task-btn'),
        addTaskMenu: document.getElementById('add-task-menu'),
        taskContainer: document.getElementById('task-container'),
        analyzeTaskTemplate: document.getElementById('analyze-task-template'),
        compareTaskTemplate: document.getElementById('compare-task-template'),
        customTaskTemplate: document.getElementById('custom-task-template'),
        costEstimate: document.getElementById('cost-estimate'),
        bulkTaskTemplate: document.getElementById("bulk-task-template"),
        autoTaskModal: document.getElementById("auto-generate-modal"),
        confirmAutoTaskBtn: document.getElementById("confirm-auto-task-btn"),
        cancelAutoTaskBtn: document.getElementById("cancel-auto-task-btn"),
        tokenEstimate: document.getElementById('token-estimate'),
        dryRunBtn: document.getElementById('dry-run-btn'),
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
    };

    // --- State Management ---
    let appState = {
        file: null,
        data: [],
        headers: [],
        isProcessing: false,
        runLog: null,
        analysisTasks: [],
        availableModels: { openai: [], ollama: [] },
        testMode: { isActive: false, task: null, currentIndex: 0 }
    };
    
    const MODEL_PRICING = {
        'gemini-1.5-flash-latest': { input: 0.35, output: 0.70 },
        'gpt-4o-mini': { input: 0.15, output: 0.60 },
        'default_openai': { input: 1.00, output: 3.00 }
    };
    const DEFAULT_SYSTEM_PROMPT = "You are an AI assistant whose sole task is to process spreadsheet rows exactly as the user’s analysis tasks specify. Output only the requested content in the exact format required—no greetings, acknowledgments, or extra commentary.";

    // --- Function Definitions ---

    const log = (message, type = 'SYSTEM') => {
        const p = document.createElement('p');
        p.innerHTML = `<span class="text-gray-500">${new Date().toLocaleTimeString()} [${type}]</span> ${message}`;
        if (type === 'ERROR') p.classList.add('text-red-400');
        ui.auditLog.appendChild(p);
        ui.auditLog.scrollTop = ui.auditLog.scrollHeight;
    };

    const estimateTokens = (text) => Math.ceil((text || '').length / 4);

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

    const getModelIdentifier = () => {
        const provider = ui.providerSelector.value;
        switch(provider) {
            case 'gemini': return ui.geminiModelSelector.value;
            case 'openai': return ui.openaiModelSelector.value;
            case 'ollama': return ui.ollamaModelSelector.value;
            default: return 'unknown';
        }
    };
    
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

        return prompt;
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
        const systemPromptTokens = estimateTokens(ui.systemPromptInput.value);
        
        appState.analysisTasks.forEach(task => {
            if (task.prompt && firstRow) {
                const userPrompt = buildPrompt(task, firstRow);
                totalInputTokens += (estimateTokens(userPrompt) + systemPromptTokens) * appState.data.length;
                totalOutputTokens += (parseInt(task.maxTokens, 10) || 150) * appState.data.length;
            }
        });
        
        const totalTokens = totalInputTokens + totalOutputTokens;
        const modelName = getModelIdentifier();
        const pricing = MODEL_PRICING[modelName] || MODEL_PRICING['default_openai'];
        const cost = ((totalInputTokens / 1_000_000) * pricing.input) + ((totalOutputTokens / 1_000_000) * pricing.output);

        ui.costEstimate.textContent = `$${cost.toFixed(4)}`;
        ui.tokenEstimate.textContent = `~${totalTokens.toLocaleString()} tokens`;
    };

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
        const apiKey = "";
        const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
        const finalPrompt = `${systemPrompt}\n\n---\n\n${userPrompt}`;
        const payload = { 
            contents: [{ role: "user", parts: [{ text: finalPrompt }] }],
            generationConfig: {
                maxOutputTokens: parseInt(maxTokens, 10) || 150,
            }
        };
        const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        const result = await response.json();
        if (!response.ok || result.error) throw new Error(result.error ? result.error.message : `HTTP error! status: ${response.status}`);
        if (result.candidates?.[0]?.content.parts[0].text) return { text: result.candidates[0].content.parts[0].text };
        throw new Error("No content in Gemini response.");
    }
    async function processWithOpenAI(userPrompt, systemPrompt, maxTokens) {
        const model = ui.openaiModelSelector.value;
        const apiKey = ui.openaiApiKeyInput.value;
        if (!apiKey) throw new Error("OpenAI API Key is required.");
        if (!model) throw new Error("Please fetch and select an OpenAI model.");
        const apiUrl = `https://api.openai.com/v1/chat/completions`;
        const payload = { 
            model: model, 
            messages: [{ role: "system", content: systemPrompt }, { role: "user", content: userPrompt }],
            max_tokens: parseInt(maxTokens, 10) || 150,
        };
        const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` }, body: JSON.stringify(payload) });
        const result = await response.json();
        if (!response.ok || result.error) throw new Error(result.error ? result.error.message : `HTTP error! status: ${response.status}`);
        if (result.choices?.[0]?.message.content) return { text: result.choices[0].message.content };
        throw new Error("No content in OpenAI response.");
    }
    async function processWithOllama(userPrompt, systemPrompt, maxTokens) {
        const model = ui.ollamaModelSelector.value;
        const url = ui.ollamaUrlInput.value;
        if (!model) throw new Error("Please fetch and select an Ollama model.");
        const apiUrl = new URL('/api/generate', url).href;
        const payload = { 
            model: model, 
            prompt: userPrompt, 
            system: systemPrompt, 
            stream: false,
            options: {
                num_predict: parseInt(maxTokens, 10) || 150,
            }
        };
        const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        const result = await response.json();
        if (!response.ok || result.error) throw new Error(result.error || `HTTP error! status: ${response.status}`);
        if (result.response) return { text: result.response };
        throw new Error("No content in Ollama response.");
    }

    function addAnalysisTask(type, taskData = null) {
        const taskId = taskData ? taskData.id : crypto.randomUUID();
        let newTask, template;
        const taskIndex = appState.analysisTasks.length;

        if (type === 'analyze') {
            newTask = taskData || { id: taskId, type: 'analyze', sourceColumn: '', outputColumn: `analysis_${taskIndex + 1}`, prompt: '', maxTokens: 150 };
            template = ui.analyzeTaskTemplate;
        } else if (type === 'compare') {
            newTask = taskData || { id: taskId, type: 'compare', sourceColumns: [], outputColumn: `comparison_${taskIndex + 1}`, prompt: '', maxTokens: 150 };
            template = ui.compareTaskTemplate;
        } else if (type === 'custom') {
            newTask = taskData || { id: taskId, type: 'custom', outputColumn: `custom_${taskIndex + 1}`, prompt: '', maxTokens: 150 };
            template = ui.customTaskTemplate;
        } else return;
        
        if (!taskData) {
            appState.analysisTasks.push(newTask);
        }

        const taskFragment = template.content.cloneNode(true);
        const taskCard = taskFragment.querySelector('.task-card');
        taskCard.dataset.taskId = taskId;

        taskCard.querySelector('input[data-type="outputColumn"]').value = newTask.outputColumn;
        taskCard.querySelector('textarea[data-type="prompt"]').value = newTask.prompt;
        taskCard.querySelector('input[data-type="maxTokens"]').value = newTask.maxTokens;

        ui.taskContainer.appendChild(taskFragment);
        
        if (type === 'analyze') {
            taskCard.querySelector('select[data-type="sourceColumn"]').value = newTask.sourceColumn;
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
        validateTasks();
        appState.analysisTasks = appState.analysisTasks.filter(task => task.id !== taskId);
        ui.taskContainer.querySelector(`.task-card[data-task-id="${taskId}"]`)?.remove();
        log('Removed analysis task.', 'INFO');
        updateCostEstimate();
    }

        validateTasks();
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
        validateTasks();

    function updateTaskColumnDropdowns() {
        document.querySelectorAll('.task-input[data-type="sourceColumn"]').forEach(dropdown => {
            const currentVal = dropdown.value;
            dropdown.innerHTML = '';
            if (appState.headers.length === 0) {
                dropdown.add(new Option('Upload a file first', ''));
                dropdown.disabled = true;
            } else {
                dropdown.add(new Option('Select a column...', ''));
                appState.headers.forEach(header => dropdown.add(new Option(header, header)));
                dropdown.disabled = false;
                if (appState.headers.includes(currentVal)) {
                    dropdown.value = currentVal;
                }
            }
        });
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
        validateTasks();
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
            analysisTasks: appState.analysisTasks,
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

                const { modelConfig } = profile;
                if (modelConfig) {
                    switch (profile.provider) {
                        case 'openai':
                            ui.openaiApiKeyInput.value = modelConfig.apiKey || '';
                            if(modelConfig.apiKey && modelConfig.model) {
                                await ui.fetchOpenAIModelsBtn.click();
                                setTimeout(() => { ui.openaiModelSelector.value = modelConfig.model; updateCostEstimate(); }, 500);
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
                            break;
                    }
                }

                // --- Apply analysis tasks ---
                if (profile.analysisTasks && Array.isArray(profile.analysisTasks)) {
                    profile.analysisTasks.forEach(task => addAnalysisTask(task.type, task));
                }

                log('Profile loaded successfully.', 'SUCCESS');

            } catch (error) {
                log(`Failed to load profile. Invalid JSON. ${error.message}`, 'ERROR');
            }
        };
        reader.readAsText(file);
        event.target.value = '';
    }

    // --- Event Listeners ---
async function generateBulkTask() {
    ui.autoTaskModal.classList.add('hidden');
    if (appState.headers.length === 0) { log('Upload a file first.', 'ERROR'); return; }
    setProcessingState(true);
    try {
        const headers = appState.headers;
        const rows = appState.data.slice(0, 5);
        let snippet = headers.join(',') + '\n';
        rows.forEach(r => { snippet += headers.map(h => r[h] || '').join(',') + '\n'; });
        const userPrompt = `Spreadsheet snippet:\n${snippet}\n\nProvide concise analysis instructions summarizing the sheet. Begin with an overview line like: "This spreadsheet consists of a header row with ${headers.length} columns, they are |${headers.join('|')}|." Also note which columns have data and which are empty, then suggest analysis for populated columns.`;
        const systemPrompt = 'You generate short task instructions for analyzing spreadsheets.';
        const result = await processWithApi(userPrompt, 200);
        const instructions = result.text.trim();
        addAnalysisTask('bulk', { id: crypto.randomUUID(), type: 'bulk', outputColumn: `bulk_analysis_${appState.analysisTasks.length + 1}`, prompt: instructions, maxTokens: 150 });
        log('Auto-generated task added.', 'INFO');
    } catch (error) {
        log(`Failed to auto-generate task: ${error.message}`, 'ERROR');
    } finally {
        setProcessingState(false);
    }
}
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
                updateTaskColumnDropdowns();
                updateCostEstimate();
                log(`Parsed ${appState.data.length} rows with columns: ${appState.headers.join(', ')}`, 'FILE');
            },
            error: (err) => { log(`File parsing error: ${err.message}`, 'ERROR'); }
        });
    });
    
    ui.providerSelector.addEventListener('change', () => {
        document.querySelectorAll('.config-group').forEach(el => el.classList.add('hidden-config'));
        document.getElementById(`${ui.providerSelector.value}-config`).classList.remove('hidden-config');
        updateCostEstimate();
    });
    
    [ui.geminiModelSelector, ui.openaiModelSelector, ui.ollamaModelSelector, ui.systemPromptInput].forEach(el => el.addEventListener('input', updateCostEstimate));
    
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

    ui.fetchOllamaModelsBtn.addEventListener('click', async () => {
        const url = ui.ollamaUrlInput.value;
        log('Fetching Ollama models...', 'API');
        ui.fetchOllamaModelsBtn.disabled = true;
        ui.ollamaModelSelector.disabled = true;
        ui.ollamaModelSelector.innerHTML = '<option>Fetching...</option>';
        try {
            const apiUrl = new URL('/api/tags', url).href;
            const response = await fetch(apiUrl);
            if (!response.ok) throw new Error(`Server returned status ${response.status}`);
            const result = await response.json();
            const models = result.models.sort((a, b) => a.name.localeCompare(b.name));
            appState.availableModels.ollama = models;
            populateModelSelector('ollama', models);
            log(`Successfully fetched ${models.length} local models.`, 'SUCCESS');
        } catch (error) {
            log(`Failed to fetch Ollama models. Is Ollama running at ${url}? Details: ${error.message}`, 'ERROR');
            ui.ollamaModelSelector.innerHTML = '<option>Failed to fetch</option>';
        } finally {
            ui.fetchOllamaModelsBtn.disabled = false;
        }
    });

    ui.addTaskBtn.addEventListener('click', () => ui.addTaskMenu.classList.toggle('hidden'));
    
    document.addEventListener('click', (e) => {
        if (!ui.addTaskDropdownContainer.contains(e.target)) {
            ui.addTaskMenu.classList.add('hidden');
        }
    });

    ui.addTaskMenu.addEventListener('click', (e) => {
        e.preventDefault();
        const taskType = e.target.dataset.taskType;
        if (taskType) {
            addAnalysisTask(taskType);
            ui.addTaskMenu.classList.add('hidden');
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
            updateTaskState(taskCard.dataset.taskId, dataType, e.target.value, e.target.dataset.index ? parseInt(e.target.dataset.index) : null);
        }
    });

    ui.runBtn.addEventListener('click', async () => {
        if (appState.isProcessing || appState.data.length === 0) { log('Cannot start run. No data loaded.', 'ERROR'); return; }
        if (appState.analysisTasks.length === 0) { log('Cannot start run. Please add at least one analysis task.', 'ERROR'); return; }
        if (!validateTasks()) return;
        
        setProcessingState(true);
        appState.runLog = { /* setup log object */ };
        log(`Starting pipeline with ${appState.analysisTasks.length} tasks...`, 'RUN');
        
        let processedData = JSON.parse(JSON.stringify(appState.data));
        const newHeaders = [...appState.headers];
        let totalItems = 0;
        const totalItemsToProcess = appState.data.length * appState.analysisTasks.length;

        for (const task of appState.analysisTasks) {
            if (!newHeaders.includes(task.outputColumn)) newHeaders.push(task.outputColumn);
            
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
        renderDataPreview();
        log('Pipeline finished successfully!', 'SUCCESS');
        setProcessingState(false);
        ui.downloadResultsBtn.style.display = 'inline-block';
        ui.downloadLogBtn.style.display = 'inline-block';
    });
    
    ui.downloadResultsBtn.addEventListener('click', () => {
        if (!appState.data || appState.data.length === 0) return;
        const csvContent = Papa.unparse(appState.data);
        const originalFilename = appState.file.name.replace(/\.[^/.]+$/, "");
        downloadFile(csvContent, `${originalFilename}_processed.csv`, 'text/csv;charset=utf-8;');
    });

    ui.downloadLogBtn.addEventListener('click', () => {
        if (!appState.runLog) return;
        const logContent = JSON.stringify(appState.runLog, null, 2);
        const originalFilename = appState.file.name.replace(/\.[^/.]+$/, "");
        downloadFile(logContent, `${originalFilename}_auditlog.json`, 'application/json');
    });

    ui.testModeBtn.addEventListener('click', () => {
        if (appState.data.length === 0) { log('Load a file to enter Test Mode.', 'ERROR'); return; }
        const validTasks = appState.analysisTasks.filter(t => t.outputColumn && ((t.type === 'analyze' && t.sourceColumn) || (t.type === 'compare' && t.sourceColumns.length > 0 && !t.sourceColumns.some(sc => !sc)) || t.type === 'custom'));
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
    ui.cancelAutoTaskBtn.addEventListener("click", () => ui.autoTaskModal.classList.add("hidden"));
    ui.confirmAutoTaskBtn.addEventListener("click", generateBulkTask);
    ui.testModePrevBtn.addEventListener('click', () => { if (appState.testMode.currentIndex > 0) { appState.testMode.currentIndex--; renderTestModeView(); } });
    ui.saveProfileBtn.addEventListener('click', saveProfile);
    ui.loadProfileInput.addEventListener('change', loadProfile);

    // --- Initial Load ---
    addAnalysisTask('analyze');
    validateTasks();
    ui.providerSelector.dispatchEvent(new Event('change'));
});
