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
        --page-bg: #f6f7f9;
        --surface: #ffffff;
        --text: #111827;
        --muted: #4b5563;
        --border: #d7dbe4;
        --border-strong: #c9cfdb;

        /* Accent (used sparingly) */
        --accent: #1f4e79;
        --accent-ink: #ffffff;

        --radius: 6px;
      }

      * { box-sizing: border-box; }

      body {
        margin: 0;
        font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
        color: var(--text);
        background: var(--page-bg);
        overflow-x: hidden; /* prevents any accidental horizontal spill */
      }

      .muted { color: var(--muted); }

      /* Top bar */
      .topbar{
        position: sticky;
        top: 0;
        z-index: 10;
        background: var(--surface);
        border-bottom: 1px solid var(--border);
        padding: 14px 22px;
      }
      .brand-title{
        font-weight: 700;
        letter-spacing: -0.01em;
        font-size: 14px;
      }
      .brand-subtitle{
        margin-top: 2px;
        font-size: 13px;
        color: var(--muted);
      }

      /* Layout */
      .container{
        max-width: 860px;
        margin: 0 auto;
        padding: 26px 22px 56px;
        min-width: 0;
      }
      .pagehead h1{
        margin: 0;
        font-size: 18px;
        font-weight: 700;
        letter-spacing: -0.01em;
      }
      .pagehead p{
        margin: 8px 0 0;
        font-size: 13px;
        overflow-wrap: anywhere;
      }

      /* Workflow */
      .workflow{
        margin-top: 18px;
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        overflow: hidden; /* ensures nothing can spill outside */
      }
      .section{
        padding: 14px 16px;
        border-bottom: 1px solid var(--border);
        min-width: 0;
      }
      .section:last-child{ border-bottom: none; }

      /* Continuous step rail (timeline-style) */
      .sectiongrid{
        display: grid;
        grid-template-columns: 34px 1fr;
        column-gap: 12px;
        min-width: 0;
      }
      .steprail{
        position: relative;
        display: flex;
        justify-content: center;
        min-width: 0;
      }
      .steprail::after{
        content:"";
        position: absolute;
        top: 26px;      /* starts just under the badge */
        bottom: -16px;  /* extends into next section for continuous flow */
        width: 2px;
        background: var(--border);
        left: 50%;
        transform: translateX(-50%);
      }
      .section.last .steprail::after{
        display: none; /* no line after final step */
      }

      .badge{
        width: 20px;
        height: 20px;
        border-radius: 999px;
        border: 1px solid var(--border-strong);
        display: grid;
        place-items: center;
        font-size: 12px;
        font-weight: 700;
        color: var(--text);
        background: var(--surface);
      }
      .badge.muted{ color: var(--muted); border-color: var(--border); }
      .badge.active{
        border-color: var(--accent);
        color: var(--accent);
      }

      .sectiontitle{
        font-weight: 700;
        font-size: 14px;
        margin: 0;
      }
      .sectiondesc{
        margin: 6px 0 0;
        color: var(--muted);
        font-size: 13px;
        overflow-wrap: anywhere;
      }

      /* Inputs / buttons */
      input[type="file"]{
        width: 100%;
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 9px 10px;
        background: var(--surface);
        min-width: 0;
      }
      button{
        border-radius: var(--radius);
        padding: 8px 10px;
        border: 1px solid var(--border);
        background: var(--surface);
        cursor: pointer;
        font-weight: 650;
        font-size: 13px;
      }
      button:hover{ border-color: var(--border-strong); }
      button:disabled{ opacity: 0.55; cursor: not-allowed; }

      .primary{
        border-color: var(--accent);
        background: var(--accent);
        color: var(--accent-ink);
      }

      .row{
        display: grid;
        grid-template-columns: 1fr;
        gap: 8px;
        min-width: 0;
      }
      .rowmeta{
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        align-items: center;
        font-size: 12px;
        color: var(--muted);
        min-width: 0;
      }
      .status{
        color: var(--muted);
      }
      .status.ok{ color: #0f5132; }

      /* Filename stays inside container */
      .filename{
        min-width: 0;
        max-width: 100%;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        display: inline-block;
      }

      .lockedtext{
        font-size: 12px;
        color: var(--muted);
        margin-top: 6px;
      }

      /* Spinner */
      .spinner{
        display: none;
        width: 14px;
        height: 14px;
        border: 2px solid var(--border-strong);
        border-top-color: transparent;
        border-radius: 999px;
        animation: spin 0.8s linear infinite;
      }
      @keyframes spin { to { transform: rotate(360deg); } }

      /* Pills */
      .pill{
        display: inline-flex;
        align-items: center;
        padding: 4px 10px;
        font-size: 12px;
        border-radius: 999px;
        border: 1px solid var(--border);
        background: var(--surface);
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
        overflow-wrap: anywhere; /* keeps long text in cells */
        word-break: break-word;
      }
      th { color: var(--muted); font-weight: 700; background: #fafbfc; }
      tr:hover td { background: #fafbfc; }
      
      /* Prevent header text from breaking into vertical letters */
      th {
        white-space: nowrap;
        overflow-wrap: normal;
        word-break: normal;
      }

      pre {
        white-space: pre-wrap;
        overflow-wrap: anywhere;
        word-break: break-word;
        background: #f4f5f7;
        border: 1px solid var(--border);
        padding: 12px;
        border-radius: var(--radius);
        overflow-x: auto;
        margin: 0;
        max-width: 100%;
      }

      .section-title{
        margin: 0 0 8px;
        font-size: 16px;
        letter-spacing: -0.01em;
      }
      .chips { display: flex; gap: 8px; flex-wrap: wrap; }

      details summary{
        cursor: pointer;
        color: var(--muted);
        font-size: 13px;
        user-select: none;
      }
    </style>
  </head>

  <body>
    <header class="topbar">
      <div class="brand">
        <div class="brand-title">Codebook Validator</div>
        <div class="brand-subtitle">Statebuilding After War</div>
      </div>
    </header>

    <main class="container">
      <div class="pagehead">
        <h1>Check Data</h1>
        <p class="muted">Upload PDF codebook instructions and a coded Excel file for automated validation of numerical codes, conditional logic, and references.</p>
      </div>

      <form action="/validate" method="post" enctype="multipart/form-data" id="workflowForm" class="workflow">
        <!-- STEP 1 -->
        <section class="section" id="step1">
          <div class="sectiongrid">
            <div class="steprail">
              <div class="badge active" id="badge1">1</div>
            </div>

            <div style="min-width:0;">
              <p class="sectiontitle">Upload codebook instructions</p>
              <p class="sectiondesc">Input format: PDF (.pdf).</p>

              <div class="row" style="margin-top:10px;">
                <input type="file" name="pdf" id="pdfInput" accept="application/pdf" required />
                <div class="rowmeta">
                  <span id="pdfStatus" class="status">No file uploaded</span>
                  <span id="pdfName" class="filename"></span>
                </div>
              </div>
            </div>
          </div>
        </section>

        <!-- STEP 2 -->
        <section class="section" id="step2">
          <div class="sectiongrid">
            <div class="steprail">
              <div class="badge muted" id="badge2">2</div>
            </div>

            <div style="min-width:0;">
              <p class="sectiontitle">Upload coded Excel workbook</p>
              <p class="sectiondesc">Input format: Excel (.xlsx).</p>

              <div class="row" style="margin-top:10px;">
                <input type="file" name="xlsx" id="xlsxInput" accept=".xlsx" required disabled />
                <div class="rowmeta">
                  <span id="xlsxStatus" class="status">No file uploaded</span>
                  <span id="xlsxName" class="filename"></span>
                </div>
                <div class="lockedtext" id="xlsxLock">Step 1 incomplete.</div>
              </div>
            </div>
          </div>
        </section>

        <!-- STEP 3 -->
        <section class="section last" id="step3">
          <div class="sectiongrid">
            <div class="steprail">
              <div class="badge muted" id="badge3">3</div>
            </div>

            <div style="min-width:0;">
              <p class="sectiontitle">Validate</p>
              <p class="sectiondesc">Executes rule extraction and workbook validation.</p>

              <div class="row" style="margin-top:10px;">
                <div style="display:flex; align-items:center; gap:10px; min-width:0;">
                  <button type="submit" id="runBtn" class="primary" disabled>Validate</button>
                  <span id="runSpinner" class="spinner" aria-hidden="true"></span>
                  <span id="runStatus" class="status"></span>
                </div>
                <div class="lockedtext" id="runLock">Step 2 incomplete.</div>
              </div>
            </div>
          </div>
        </section>
      </form>

      <!-- OUTPUT (always visible) -->
      <section class="workflow" style="margin-top:16px;">
        <div class="section" style="border-bottom: 1px solid var(--border);">
          <p class="sectiontitle" style="margin:0;">Validation Output</p>
          <p class="sectiondesc" style="margin-bottom:0;">Results display after a validation run completes.</p>
        </div>

        <div class="section">
          {% if report %}
            <div class="section-title">Summary</div>

            <div class="chips" style="margin-bottom:12px;">
              <span class="pill error">Errors: {{ report.workbook_summary.total_errors }}</span>
              <span class="pill warn">Warnings: {{ report.workbook_summary.total_warnings }}</span>
              <span class="pill ok">Tabs checked: {{ report.workbook_summary.tabs_validated }}</span>
            </div>

            <div style="display:grid; gap:6px; font-size: 13px; min-width:0;">
              <div><span class="muted"><b>PDF variables detected:</b></span> {{ report.pdf_summary.variables_detected }}</div>
              <div><span class="muted"><b>Candidate rules extracted:</b></span> {{ report.pdf_summary.rules_extracted }}</div>
              <div><span class="muted"><b>Rules by confidence:</b></span> {{ report.pdf_summary.rules_by_confidence }}</div>
            </div>

            {% for tab in report.tabs %}
              <div style="margin-top:16px; min-width:0;">
                <div class="section-title" style="margin-bottom:6px;">Tab: {{ tab.tab_name }}</div>
                <div class="chips" style="margin: 10px 0 14px;">
                  <span class="pill error">Errors: {{ tab.summary.errors }}</span>
                  <span class="pill warn">Warnings: {{ tab.summary.warnings }}</span>
                  <span class="pill ok">OK checks: {{ tab.summary.ok }}</span>
                </div>

                <div class="section-title" style="margin-top:0;">Issues</div>

                {% if tab.issues|length == 0 %}
                  <p class="muted" style="margin:0; font-size:13px;">No issues found.</p>
                {% else %}
                  <div style="overflow-x:auto; max-width:100%;">
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

            <div style="margin-top: 14px;">
              <details>
                <summary>Raw JSON</summary>
                <div style="margin-top:10px;"><pre>{{ report_json }}</pre></div>
              </details>
            </div>
          {% else %}
            <p class="muted" style="margin:0; font-size: 13px;">No report generated.</p>
          {% endif %}
        </div>
      </section>
    </main>

    <script>
      const pdfInput = document.getElementById("pdfInput");
      const xlsxInput = document.getElementById("xlsxInput");
      const runBtn = document.getElementById("runBtn");

      const badge2 = document.getElementById("badge2");
      const badge3 = document.getElementById("badge3");
      const pdfStatus = document.getElementById("pdfStatus");
      const xlsxStatus = document.getElementById("xlsxStatus");
      const pdfName = document.getElementById("pdfName");
      const xlsxName = document.getElementById("xlsxName");
      const xlsxLock = document.getElementById("xlsxLock");
      const runLock = document.getElementById("runLock");

      const runSpinner = document.getElementById("runSpinner");
      const runStatus = document.getElementById("runStatus");
      const form = document.getElementById("workflowForm");

      function update() {
        const hasPdf = pdfInput.files && pdfInput.files.length > 0;
        const hasXlsx = xlsxInput.files && xlsxInput.files.length > 0;

        // Step 1
        if (hasPdf) {
          pdfStatus.textContent = "File uploaded";
          pdfStatus.classList.add("ok");
          pdfName.textContent = pdfInput.files[0].name;
        } else {
          pdfStatus.textContent = "No file uploaded";
          pdfStatus.classList.remove("ok");
          pdfName.textContent = "";
        }

        // Step 2 unlock
        xlsxInput.disabled = !hasPdf;
        badge2.classList.toggle("muted", !hasPdf);
        badge2.classList.toggle("active", hasPdf);

        xlsxLock.textContent = hasPdf ? "" : "Step 1 incomplete.";
        xlsxLock.style.display = hasPdf ? "none" : "block";

        if (hasXlsx) {
          xlsxStatus.textContent = "File uploaded";
          xlsxStatus.classList.add("ok");
          xlsxName.textContent = xlsxInput.files[0].name;
        } else {
          xlsxStatus.textContent = "No file uploaded";
          xlsxStatus.classList.remove("ok");
          xlsxName.textContent = "";
        }

        // Step 3 unlock
        const canRun = hasPdf && hasXlsx;
        runBtn.disabled = !canRun;

        badge3.classList.toggle("muted", !canRun);
        badge3.classList.toggle("active", canRun);

        runLock.textContent = canRun ? "" : "Step 2 incomplete.";
        runLock.style.display = canRun ? "none" : "block";
      }

      pdfInput.addEventListener("change", () => {
        // PDF changes invalidate later inputs
        xlsxInput.value = "";
        update();
      });

      xlsxInput.addEventListener("change", () => {
        update();
      });

      form.addEventListener("submit", () => {
        runSpinner.style.display = "inline-block";
        runStatus.textContent = "Validating";
        runBtn.disabled = true;
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