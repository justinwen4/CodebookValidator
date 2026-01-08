"""
Codebook Validator — MVP v1 (1-week build)

This is a small Flask web app that:
1) Accepts a codebook PDF upload
2) Accepts an Excel codebook (.xlsx) upload (multi-tab)
3) Extracts conservative "candidate rules" from the PDF (heuristics)
4) Validates the Excel against those rules + structural checks
5) Returns an explainable Errors/Warnings report

Important MVP constraints (from PRD):
- Rows represent variables, not cases.
- Narrative columns should not be validated.
- Blank vs "NA" have different meanings.
"""

from __future__ import annotations

import os
import uuid
from typing import Dict, Any

from flask import Flask, request, render_template_string, jsonify

from validator import (
    extract_pdf_model,
    parse_excel_workbook,
    validate_workbook,
)

APP = Flask(__name__)

# Store uploads in a local folder so the app can read them.
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


HTML = """
<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <title>Codebook Validator (MVP)</title>
    <style>
      body { font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto; margin: 2rem; }
      .card { border: 1px solid #ddd; border-radius: 12px; padding: 1rem 1.25rem; margin-bottom: 1rem; }
      .muted { color: #666; }
      label { display: inline-block; margin-top: 0.5rem; font-weight: 600; }
      input[type=file] { margin-top: 0.35rem; }
      button { margin-top: 0.8rem; padding: 0.55rem 0.9rem; border-radius: 10px; border: 1px solid #444; background: #111; color: white; cursor: pointer; }
      button:hover { opacity: 0.9; }
      table { border-collapse: collapse; width: 100%; }
      th, td { border: 1px solid #ddd; padding: 0.5rem; text-align: left; vertical-align: top; }
      th { background: #f6f6f6; }
      .pill { display: inline-block; padding: 0.1rem 0.5rem; border-radius: 999px; font-size: 0.85rem; }
      .error { background: #ffe8e8; border: 1px solid #ffb0b0; }
      .warn { background: #fff7db; border: 1px solid #ffe08a; }
      .ok { background: #e8fff0; border: 1px solid #a9f0c3; }
      pre { white-space: pre-wrap; background: #f7f7f7; padding: 0.75rem; border-radius: 10px; overflow-x: auto; }
    </style>
  </head>
  <body>
    <h1>Codebook Validator (MVP)</h1>
    <p class="muted">
      Upload a codebook PDF + a coded Excel workbook. This app returns errors/warnings with explanations.
      It validates the <b>numerical code</b> and some simple conditionals; it does <b>not</b> validate narratives.
    </p>

    <div class="card">
      <h2>1) Upload</h2>
      <form action="/validate" method="post" enctype="multipart/form-data">
        <label>Codebook PDF</label><br/>
        <input type="file" name="pdf" accept="application/pdf" required/><br/>

        <label>Excel codebook (.xlsx)</label><br/>
        <input type="file" name="xlsx" accept=".xlsx" required/><br/>

        <button type="submit">Validate</button>
      </form>
    </div>

    {% if report %}
      <div class="card">
        <h2>2) Summary</h2>
        <p><b>PDF variables detected:</b> {{ report.pdf_summary.variables_detected }}</p>
        <p><b>Candidate rules extracted:</b> {{ report.pdf_summary.rules_extracted }}</p>
        <p><b>Rules by confidence:</b> {{ report.pdf_summary.rules_by_confidence }}</p>
        <p><b>Workbook tabs validated:</b> {{ report.workbook_summary.tabs_validated }}</p>
        <p><b>Total issues:</b>
          <span class="pill error">Errors: {{ report.workbook_summary.total_errors }}</span>
          <span class="pill warn">Warnings: {{ report.workbook_summary.total_warnings }}</span>
        </p>
      </div>

      {% for tab in report.tabs %}
        <div class="card">
          <h2>Tab: {{ tab.tab_name }}</h2>
          <p>
            <span class="pill error">Errors: {{ tab.summary.errors }}</span>
            <span class="pill warn">Warnings: {{ tab.summary.warnings }}</span>
            <span class="pill ok">OK checks: {{ tab.summary.ok }}</span>
          </p>

          <h3>Issues</h3>
          {% if tab.issues|length == 0 %}
            <p class="muted">No issues found.</p>
          {% else %}
            <table>
              <thead>
                <tr>
                  <th>Severity</th>
                  <th>Variable</th>
                  <th>Message</th>
                  <th>Expected</th>
                  <th>Actual</th>
                  <th>Evidence</th>
                </tr>
              </thead>
              <tbody>
                {% for issue in tab.issues %}
                  <tr>
                    <td>
                      {% if issue.severity == "error" %}
                        <span class="pill error">error</span>
                      {% else %}
                        <span class="pill warn">warning</span>
                      {% endif %}
                    </td>
                    <td>{{ issue.variable }}</td>
                    <td>{{ issue.message }}</td>
                    <td>{{ issue.expected }}</td>
                    <td>{{ issue.actual }}</td>
                    <td><pre>{{ issue.evidence }}</pre></td>
                  </tr>
                {% endfor %}
              </tbody>
            </table>
          {% endif %}
        </div>
      {% endfor %}

      <div class="card">
        <h2>3) Raw JSON (for debugging)</h2>
        <pre>{{ report_json }}</pre>
      </div>
    {% endif %}
  </body>
</html>
"""


@APP.get("/")
def home():
    """Render the upload form."""
    return render_template_string(HTML, report=None)


@APP.post("/validate")
def validate_route():
    """
    Main route:
    - saves the uploaded files
    - extracts a PDF "model" (variables + candidate rules)
    - parses the Excel workbook into a normalized structure
    - validates and returns a report
    """
    pdf_file = request.files.get("pdf")
    xlsx_file = request.files.get("xlsx")

    if not pdf_file or not xlsx_file:
        return "Missing pdf or xlsx upload", 400

    pdf_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}.pdf")
    xlsx_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex}.xlsx")
    pdf_file.save(pdf_path)
    xlsx_file.save(xlsx_path)

    pdf_model = extract_pdf_model(pdf_path)
    workbook = parse_excel_workbook(xlsx_path)
    report = validate_workbook(pdf_model, workbook)

    # Flask/Jinja can render dicts, but we also want raw JSON for easy debugging.
    report_dict: Dict[str, Any] = report
    report_json = jsonify(report_dict).get_data(as_text=True)

    return render_template_string(HTML, report=report_dict, report_json=report_json)


@APP.get("/api/validate")
def api_help():
    """Small helper message so people don't hit /api/validate accidentally."""
    return jsonify({
        "message": "Use POST /validate with multipart form-data fields: pdf, xlsx"
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5050"))
    APP.run(host="0.0.0.0", port=port, debug=True)