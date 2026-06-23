# DOCX Formatting Standards Validator 
## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`

## Python Local Tool
This folder contains a complete, working Python-based implementation of the DOCX Formatting Standards Validator. It features a Flask backend and a modern web dashboard.

---

## 📂 Folder Structure

```
python_app/
├── app.py             # Main Flask server + XML-level paragraph parser
├── requirements.txt   # Python package dependencies
└── README.md          # Local setup & running instructions (this file)
```

---

## ⚙️ Prerequisites \& Setup

Make sure you have **Python 3.8 or newer** installed on your system.

### 1. Install Dependencies

Open your terminal or command prompt inside the `python_app` folder, and use `pip` to install the required packages:

```bash
pip install -r requirements.txt
```

Alternatively, install them manually:
```bash
pip install Flask python-docx lxml
```

---

## 🚀 Running the Local Web UI

1. Run the Flask application:
   ```bash
   python app.py
   ```
   or
   ```bash
   py app.py
   ```

3. Open your web browser and navigate to:
   ```
   http://127.0.0.1:5000
   ```

4. Enter the paths:
   - **Target Scanning Directory**: The absolute path to the directory containing your `.docx` files.
   - **Output Directory**: (Optional) Path to save the extracted JSON and CSV reports.

5. Click **Run Analysis** to scan and validate documents instantly.

---

## ⚡ Standards \& Validation Rules Applied

Each paragraph of your documents is extracted down to the word level using `python-docx` and XML parsing. The rules enforced are:

| Level ID | Expected Level | Font Name | Font Size | Bold Status | Expected Alignment | Expected Indent |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `heading1` | Level 1 Heading | `Human Black Condensed` | `20 pt` | `True` | `LEFT` | `0 pt` |
| `heading2` | Level 2 Heading | `Human Black Condensed` | `11.5 pt` | `True` | `LEFT` | `0 pt` |
| `heading3` | Level 3 Heading | `Human Black Condensed` | `10.5 pt` | `True` | `LEFT` | `0 pt` |
| `normal` | Normal Text | `Human Light Condensed` | `10.5 pt` | `False` | `LEFT` | `0 pt` |
| `boldText` | Text Level Bold | `Human Bold Condensed` | `10.5 pt` | `True` | `LEFT` | `0 pt` |
| `formName` | Form Name (Footer/End) | `Human Light Condensed` | `7 pt` | `False` | `LEFT` | `0 pt` |

---

## 💾 Exported Outputs

If you specify an output folder path in the UI, the script will automatically write:
1. `formatting_analysis_report.json`: Full detailed report of all file paragraphs, including run-level properties.
2. `formatting_analysis_report.csv`: Flat, spreadsheet-ready table listing paragraph texts, dominant formatting, expected levels, status (Pass/Fail) and specific issue highlights.
