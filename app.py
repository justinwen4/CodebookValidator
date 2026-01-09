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
    <title>Codebook Validator</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      :root{
        --bg: #f7f8fb;
        --card: #ffffff;
        --text: #0b0f19;
        --muted: #5b6474;
        --border: #e6e8ee;
        --shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
        --radius: 16px;
      }

      * { box-sizing: border-box; }

      body {
        margin: 0;
        font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
        color: var(--text);
        background: var(--bg);
      }

      .muted { color: var(--muted); }

      /* Top bar */
      .topbar{
        position: sticky;
        top: 0;
        z-index: 10;
        background: rgba(255,255,255,0.92);
        backdrop-filter: blur(10px);
        border-bottom: 1px solid var(--border);
        padding: 18px 22px;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .brand-title{
        font-weight: 800;
        letter-spacing: -0.02em;
      }
      .brand-subtitle{
        margin-top: 2px;
        font-size: 13px;
        color: var(--muted);
      }
      .nav a{
        text-decoration: none;
        color: var(--muted);
        font-size: 14px;
        margin-left: 16px;
      }
      .nav a:hover{ color: var(--text); }

      /* Layout */
      .container{
        max-width: 1100px;
        margin: 0 auto;
        padding: 26px 22px 56px;
      }
      .pagehead h1{
        margin: 0;
        font-size: 28px;
        letter-spacing: -0.02em;
      }
      .pagehead p{ margin: 10px 0 0; }

      /* Cards */
      .card{
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        box-shadow: var(--shadow);
        padding: 18px;
        margin-top: 16px;
      }

      /* Step wizard */
      .steps {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 14px;
        margin-top: 14px;
      }
      @media (max-width: 980px) {
        .steps { grid-template-columns: 1fr; }
      }
      .stepcard{
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 16px;
        background: #fbfcff;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
        display: flex;
        gap: 12px;
        min-height: 210px;
      }
      .stepcard.locked{ opacity: 0.55; }
      .stepnum{
        width: 38px;
        height: 38px;
        border-radius: 12px;
        border: 1px solid var(--border);
        background: #fff;
        display: grid;
        place-items: center;
        font-weight: 800;
      }
      .stepbody{ width: 100%; }
      .steptitle{ font-weight: 800; margin-bottom: 6px; }
      .stephelp{ margin-bottom: 10px; }

      /* Inputs / buttons */
      input[type="file"]{
        width: 100%;
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 10px 12px;
        background: #fff;
      }
      button{
        border-radius: 12px;
        padding: 10px 14px;
        border: 1px solid var(--border);
        background: #fff;
        cursor: pointer;
        font-weight: 650;
      }
      button:hover{ border-color: #cfd6e5; }
      button:disabled{ opacity: 0.55; cursor: not-allowed; }
      .primary{
        border-color: #0b0f19;
        background: #0b0f19;
        color: #fff;
      }
      .primary:hover{ filter: brightness(0.97); }

      /* Meta row */
      .stepmeta{
        margin-top: 10px;
        display: flex;
        gap: 10px;
        align-items: center;
        flex-wrap: wrap;
        font-size: 13px;
      }
      .link{
        color: var(--muted);
        text-decoration: none;
      }
      .link:hover{
        color: var(--text);
        text-decoration: underline;
      }

      /* Pills / badges */
      .pill{
        display: inline-flex;
        align-items: center;
        padding: 4px 10px;
        font-size: 12px;
        border-radius: 999px;
        border: 1px solid var(--border);
        background: #fff;
        color: var(--muted);
      }
      .pill.ok{ background: #eafff2; border-color: #b8f0c9; color: #1b6b3a; }
      .pill.warn{ background: #fff4da; border-color: #ffe2a1; color: #7a5a00; }
      .pill.error{ background: #ffe7e7; border-color: #ffb3b3; color: #8a1f1f; }

      /* Tables */
      table { border-collapse: collapse; width: 100%; }
      th, td {
        border-bottom: 1px solid var(--border);
        padding: 10px;
        text-align: left;
        vertical-align: top;
        font-size: 14px;
      }
      th { color: var(--muted); font-weight: 700; background: #fbfcff; }
      tr:hover td { background: #fbfcff; }

      pre {
        white-space: pre-wrap;
        background: #f6f7fb;
        border: 1px solid var(--border);
        padding: 12px;
        border-radius: 14px;
        overflow-x: auto;
        margin: 0;
      }

      /* Results spacing */
      .section-title{
        margin: 0 0 8px;
        font-size: 16px;
        letter-spacing: -0.01em;
      }
      .chips { display: flex; gap: 8px; flex-wrap: wrap; }
    </style>
  </head>

  <body>
    <header class="topbar">
      <div class="brand">
        <div class="brand-title">Codebook Validator</div>
        <div class="brand-subtitle">Lab-grade dataset checks</div>
      </div>
      <nav class="nav">
        <a href="#">Help</a>
        <a href="#">How it works</a>
        <a href="#">Sample files</a>
      </nav>
    </header>

    <main class="container">
      <div class="pagehead">
        <h1>Check Data</h1>
        <p class="muted">
          Upload your codebook instructions and coded Excel file to receive validation feedback.
          We validate <b>numerical codes</b> and simple conditionals; we do <b>not</b> validate narratives.
        </p>
      </div>

      <div class="card">
        <div class="section-title">Validate Codebook</div>
        <p class="muted" style="margin:0;">Complete each step to unlock the next. Validation runs once you submit Step 3.</p>

        <form action="/validate" method="post" enctype="multipart/form-data" id="wizardForm">
          <div class="steps">

            <!-- STEP 1 -->
            <div class="stepcard" id="step1">
              <div class="stepnum">1</div>
              <div class="stepbody">
                <div class="steptitle">Upload Codebook Instructions (PDF)</div>
                <div class="stephelp muted">Select the instruction manual PDF.</div>

                <input type="file" name="pdf" id="pdfInput" accept="application/pdf" required />

                <div class="stepmeta">
                  <span id="pdfStatus" class="pill">Not selected</span>
                  <span id="pdfName" class="muted"></span>
                  <a href="#" id="pdfReplace" class="link">Replace</a>
                </div>
              </div>
            </div>

            <!-- STEP 2 -->
            <div class="stepcard locked" id="step2">
              <div class="stepnum">2</div>
              <div class="stepbody">
                <div class="steptitle">Upload Coded Excel File (.xlsx)</div>
                <div class="stephelp muted" id="xlsxHelp">Locked until Step 1 is complete.</div>

                <input type="file" name="xlsx" id="xlsxInput" accept=".xlsx" required disabled />

                <div class="stepmeta">
                  <span id="xlsxStatus" class="pill">Not selected</span>
                  <span id="xlsxName" class="muted"></span>
                  <a href="#" id="xlsxReplace" class="link">Replace</a>
                </div>
              </div>
            </div>

            <!-- STEP 3 -->
            <div class="stepcard locked" id="step3">
              <div class="stepnum">3</div>
              <div class="stepbody">
                <div class="steptitle">Run Validation</div>
                <div class="stephelp muted" id="runHelp">Locked until Step 2 is complete.</div>

                <button type="submit" id="runBtn" class="primary" disabled>Validate</button>

                <div class="stepmeta muted">
                  Results will appear below after the run finishes.
                </div>
              </div>
            </div>

          </div>
        </form>
      </div>

      {% if report %}
        <div class="card">
          <div class="section-title">Summary</div>

          <div class="chips" style="margin-bottom:12px;">
            <span class="pill error">Errors: {{ report.workbook_summary.total_errors }}</span>
            <span class="pill warn">Warnings: {{ report.workbook_summary.total_warnings }}</span>
            <span class="pill ok">Tabs checked: {{ report.workbook_summary.tabs_validated }}</span>
          </div>

          <div style="display:grid; gap:6px;">
            <div><span class="muted"><b>PDF variables detected:</b></span> {{ report.pdf_summary.variables_detected }}</div>
            <div><span class="muted"><b>Candidate rules extracted:</b></span> {{ report.pdf_summary.rules_extracted }}</div>
            <div><span class="muted"><b>Rules by confidence:</b></span> {{ report.pdf_summary.rules_by_confidence }}</div>
          </div>
        </div>

        {% for tab in report.tabs %}
          <div class="card">
            <div class="section-title">Tab: {{ tab.tab_name }}</div>
            <div class="chips" style="margin: 10px 0 14px;">
              <span class="pill error">Errors: {{ tab.summary.errors }}</span>
              <span class="pill warn">Warnings: {{ tab.summary.warnings }}</span>
              <span class="pill ok">OK checks: {{ tab.summary.ok }}</span>
            </div>

            <div class="section-title" style="margin-top:0;">Issues</div>

            {% if tab.issues|length == 0 %}
              <p class="muted" style="margin:0;">No issues found.</p>
            {% else %}
              <div style="overflow-x:auto;">
                <table>
                  <thead>
                    <tr>
                      <th style="width:110px;">Severity</th>
                      <th style="width:190px;">Variable</th>
                      <th>Message</th>
                      <th style="width:220px;">Expected</th>
                      <th style="width:220px;">Actual</th>
                      <th style="width:340px;">Evidence</th>
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
                        <td><b>{{ issue.variable }}</b></td>
                        <td>{{ issue.message }}</td>
                        <td class="muted">{{ issue.expected }}</td>
                        <td class="muted">{{ issue.actual }}</td>
                        <td><pre>{{ issue.evidence }}</pre></td>
                      </tr>
                    {% endfor %}
                  </tbody>
                </table>
              </div>
            {% endif %}
          </div>
        {% endfor %}

        <div class="card">
          <div class="section-title">Raw JSON (debug)</div>
          <pre>{{ report_json }}</pre>
        </div>
      {% endif %}
    </main>

    <script>
      const pdfInput = document.getElementById("pdfInput");
      const xlsxInput = document.getElementById("xlsxInput");
      const runBtn = document.getElementById("runBtn");

      const step2 = document.getElementById("step2");
      const step3 = document.getElementById("step3");

      const pdfStatus = document.getElementById("pdfStatus");
      const xlsxStatus = document.getElementById("xlsxStatus");
      const pdfName = document.getElementById("pdfName");
      const xlsxName = document.getElementById("xlsxName");

      const xlsxHelp = document.getElementById("xlsxHelp");
      const runHelp = document.getElementById("runHelp");

      function setLocked(el, locked) {
        el.classList.toggle("locked", locked);
      }

      function resetFrom(step) {
        if (step <= 1) {
          // Changing PDF invalidates later steps (UI-only)
          xlsxInput.value = "";
          xlsxInput.disabled = true;
          setLocked(step2, true);
          xlsxStatus.textContent = "Not selected";
          xlsxStatus.className = "pill";
          xlsxName.textContent = "";
          xlsxHelp.textContent = "Locked until Step 1 is complete.";
        }
        if (step <= 2) {
          runBtn.disabled = true;
          setLocked(step3, true);
          runHelp.textContent = "Locked until Step 2 is complete.";
        }
      }

      function update() {
        const hasPdf = pdfInput.files && pdfInput.files.length > 0;
        const hasXlsx = xlsxInput.files && xlsxInput.files.length > 0;

        // Step 1 status
        if (hasPdf) {
          pdfStatus.textContent = "Selected";
          pdfStatus.className = "pill ok";
          pdfName.textContent = pdfInput.files[0].name;
        } else {
          pdfStatus.textContent = "Not selected";
          pdfStatus.className = "pill";
          pdfName.textContent = "";
        }

        // Step 2 unlock
        xlsxInput.disabled = !hasPdf;
        setLocked(step2, !hasPdf);
        xlsxHelp.textContent = hasPdf
          ? "Select the coded Excel workbook (.xlsx)."
          : "Locked until Step 1 is complete.";

        // Step 2 status
        if (hasXlsx) {
          xlsxStatus.textContent = "Selected";
          xlsxStatus.className = "pill ok";
          xlsxName.textContent = xlsxInput.files[0].name;
        } else {
          xlsxStatus.textContent = "Not selected";
          xlsxStatus.className = "pill";
          xlsxName.textContent = "";
        }

        // Step 3 unlock
        const canRun = hasPdf && hasXlsx;
        runBtn.disabled = !canRun;
        setLocked(step3, !canRun);
        runHelp.textContent = canRun
          ? "Ready to validate."
          : "Locked until Step 2 is complete.";
      }

      pdfInput.addEventListener("change", () => {
        resetFrom(1);
        update();
      });

      xlsxInput.addEventListener("change", () => {
        resetFrom(2);
        update();
      });

      document.getElementById("pdfReplace").addEventListener("click", (e) => {
        e.preventDefault();
        pdfInput.value = "";
        resetFrom(1);
        update();
      });

      document.getElementById("xlsxReplace").addEventListener("click", (e) => {
        e.preventDefault();
        xlsxInput.value = "";
        resetFrom(2);
        update();
      });

      update();
    </script>
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