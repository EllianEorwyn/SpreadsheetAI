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

### 🔎 Validation Tools
- Reliability Checks (Cohen's Kappa and exact match)
- Consistency between columns (Pearson or pairwise)
- Missing Output Audit
- Prompt Stability Test (token-level diff)
- Face Validity Export (sampled rows with human vs. AI output)

### 🛠️ Model Configuration
- Supports OpenAI, Gemini, and local Ollama models
- Dynamic model fetch for OpenAI and Ollama
- Cost and token estimates updated live
- Auto-fill models based on API key / local instance

### 💾 Profile Management
- Save and load complete analysis profiles as JSON
- Includes model config, system prompt, tasks, and validation settings

### 📥 Export Options
- Download processed CSV
- Export full audit log (token usage, prompts, outputs)
- Export face validity samples

