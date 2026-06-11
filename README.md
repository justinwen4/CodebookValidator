# Codebook Validator

A small Flask web app that validates an Excel codebook workbook against rules extracted from a codebook PDF.

## What this does
- Upload a **codebook PDF**
- Upload a coded **Excel workbook (.xlsx)** with multiple tabs
- The app:
  - Extracts variable IDs like `[educ_post_expand]` from the PDF
  - Extracts conservative “candidate rules” from the PDF (heuristics)
  - Parses the Excel workbook into a normalized structure
  - Validates **numerical codes** (does **not** validate narrative text)
  - Displays an explainable Errors/Warnings report per tab

## Project structure
- `app.py`
  - Flask server + simple HTML UI (upload form + results)
  - Calls the functions in `validator.py`
- `validator.py`
  - PDF parsing
  - Rule extraction heuristics
  - Excel parsing (openpyxl)
  - Validation engine (produces explainable issues)
- `requirements.txt`
  - Python dependencies

## Setup (macOS / zsh)

### 1) Go to the project folder
```bash
cd /Users/overp/Python/CodebookValidator
