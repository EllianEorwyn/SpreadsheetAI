# SpreadsheetAI

An LLM-Enhanced Spreadsheet Classifier App that lets you upload a spreadsheet, define custom analyses, and batch-process each row through an LLM—all in your browser, with live token-count and cost estimates.

## 🚀 Purpose & Scope

- **Goal:**  
  1. Upload a CSV/XLSX spreadsheet  
  2. Define one or more “analysis tasks” (e.g. analyze a column, compare columns)  
  3. Let the LLM process each row  
  4. Download a new spreadsheet with your analysis results  

- **Core Values:**  
  - **Transparency:** see exactly what prompts are sent  
  - **Reproducibility:** save/load profiles & pipelines  
  - **User Control:** pick models, tune tokens & cost  
  - **Privacy:** all processing happens client-side  
  - **Cost-Awareness:** live token and dollar estimates  

---

## 🎨 High-Level Workflow & UI Elements

1. **File Upload**  
   - **Component:** **Upload Spreadsheet** (CSV/XLSX drag-and-drop or click)  
   - **Outcome:** Data Preview Table shows your header row + first 5 rows  

2. **Analysis Pipeline**  
   - **Panel:** **Analysis Pipeline** (collapsible)  
   - **Action:** **Add Task** dropdown → “Analyze Column” / “Compare Columns”  
   - **Task Card:** choose source column(s), name output column, enter instructions, set max tokens  

3. **Global Settings**  
   - **Sidebar:** **Settings**  
   - **Controls:** API provider, API key, model selector, cost per million tokens, local-LLM toggle  

4. **Token & Cost Estimator**  
   - **Estimator Panel:** shows input tokens, estimated output tokens, estimated cost  
   - **Controls:** “Recalculate” button, “Dry Run” toggle  

5. **Processing Controls**  
   - **Buttons:**  
     - **Run Batch Analysis**  
     - **Enable Test Mode** toggle  
     - **Start Interactive Review** (row-by-row QA)  

6. **Progress & Output**  
   - **Progress Bar**  
   - **Status Toasts** (errors, warnings)  
   - **Download Processed Spreadsheet** button
   - **View Audit Log** (JSON/CSV)

### Prompt Variables

Task instructions can reference spreadsheet data using placeholder syntax:

- `{{COLUMN}}` – replaced with the value from the selected source column (for Analyze tasks).
- `{{Column Name}}` – replaced with the value from any column in the current row.

Each prompt is also prepended with a hidden "row context" block containing all column values so the model has the full row data without repeating it in the output.

---

## 💻 Running Locally

1. Clone or download this repository.
2. Open `index.html` in your preferred browser to launch the app offline. No server is required.

### Optional npm tooling

If network access is available, install optional development tools:

```bash
npm install
```

This installs potential linters, bundlers, or other dev utilities.

## 🤝 Contributing

We welcome issues and pull requests.

1. Fork the project and create a feature branch.
2. Make your changes with clear commit messages.
3. Open a PR referencing any related issues.
4. After review, your contribution can be merged.

