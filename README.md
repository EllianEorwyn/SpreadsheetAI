# SpreadsheetAI

An LLM-Enhanced Spreadsheet Classifier App that lets you upload a spreadsheet, define custom analyses, and batch-process each row through an LLM—all in your browser, with live token-count and cost estimates.

## 🔍 Key Features

### 🧠 LLM-Driven Spreadsheet Processing
- Upload a CSV and define custom NLP or logic tasks.
- Model processes each row individually based on user-defined prompts.
- Results are written back to the spreadsheet in new or specified columns.

### 🧩 Multi-Task Architecture
- Supports multiple task types:
  - **Analyze**: Single-column analysis (e.g., classify, summarize).
  - **Compare**: Cross-column comparisons (e.g., detect contradiction, compare stance).
  - **Custom**: Fully flexible prompts referencing any column.
  - **Auto-Generated**: Uses the model to suggest tasks for empty columns.

### 📋 Built-in Task Presets
- One-click templates for common social-science and text analysis tasks:
  - Sentiment, Emotion, Topic Detection
  - Summarization
  - Named Entity Extraction
  - Headline Generation
  - Stance Detection
  - Contradiction Detection

### ✨ Prompt Engineering Support
- Use `{{ColumnName}}` to dynamically insert row data.
- Reference full rows using the `Include row context` toggle.
- Auto-generate task prompts based on data preview and task goal.

### 🔍 Test Mode
- Step through rows manually for any task.
- Preview the row, generated prompt, and AI response.
- Rerun and edit tasks interactively.

### 📊 Inline Output Analysis
- For each generated output column:
  - Show type detection (numeric, categorical, text)
  - Display charts:
    - Histogram (numeric)
    - Bar chart (categorical)
  - Word cloud (text)
  - Export summary statistics and frequency tables
  - Log scale toggle for numeric data
  - Manual override to treat columns as text or categorical

### 🔎 Validation Tools
- Reliability Checks (Cohen's Kappa, Krippendorff's Alpha, and exact match) with a dedicated **Intercoder Reliability** dashboard section
- Consistency between columns (Pearson or pairwise)
- Missing Output Audit
- Prompt Stability Test (token-level diff)
- Face Validity Export (sampled rows with human vs. AI output)

### 🖥️ AI Output Analysis Dashboard
- Explore charts and summary statistics for each generated column
- The **Intercoder Reliability** panel visualizes agreement scores
- Krippendorff's Alpha supports multi-label coding by treating coder annotations as sets and calculating disagreement via Jaccard distance
- Observed disagreement averages pairwise distances across units, while expected disagreement samples from the empirical distribution of code sets
- Partial overlap earns a distance between 0 and 1, so partial agreement is rewarded more than total disagreement but less than a perfect match
- Alpha weighting can be switched between **Basic** (binary match), **Jaccard**, or **MASI** distance
- The reliability panel reports the observed disagreement ($D_o$) and expected disagreement ($D_e$) used in the alpha calculation
- $\alpha = 1 - \frac{D_o}{D_e}$ where each distance function returns a value in $[0,1]$

### 🛠️ Model Configuration
- Supports OpenAI, Gemini, and local Ollama models
- Dynamic model fetch for OpenAI and Ollama
- Cost and token estimates updated live
- Auto-fill models based on API key / local instance

## 🛠️ Using Local Ollama Models

Newer versions of Ollama expose an OpenAI-compatible model list at `/v1/models` while
older releases use the legacy `/api/tags` route. The app now tries both endpoints
automatically.

If the browser console shows CORS errors (often a 403 on the preflight `OPTIONS` call),
start Ollama with your front-end origin whitelisted:

```
export OLLAMA_ORIGINS=http://localhost:3000,http://127.0.0.1:3000
# If accessing from Docker/WSL/LAN also bind the server:
export OLLAMA_HOST=0.0.0.0:11434
ollama serve
```

You can verify the server responds to both endpoints:

```
curl -i http://127.0.0.1:11434/v1/models
curl -i http://127.0.0.1:11434/api/tags
curl -i -X OPTIONS http://127.0.0.1:11434/v1/models \
  -H "Origin: http://localhost:3000" \
  -H "Access-Control-Request-Method: GET"
```


### 💾 Profile Management
- Save and load complete analysis profiles as JSON
- Includes model config, system prompt, tasks, and validation settings

### 📥 Export Options
- Download processed CSV
- Export full audit log (token usage, prompts, outputs)
- Export face validity samples

