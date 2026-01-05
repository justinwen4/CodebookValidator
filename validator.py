"""
validator.py

All the "real logic" lives here so app.py stays very small.

This module contains:
- PDF parsing + rule extraction (conservative heuristics)
- Excel parsing (openpyxl) into a normalized structure
- Validation engine that produces explainable issues

We keep the MVP simple and conservative:
- Only validate numerical-code column (NOT narrative text).
- Treat unclear / low-confidence rules as warnings-only.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple, Any

import openpyxl

# PDF parsing: try pdfplumber if installed (best), else fall back to PyPDF2.
try:
    import pdfplumber  # type: ignore
except Exception:
    pdfplumber = None

try:
    from PyPDF2 import PdfReader  # type: ignore
except Exception:
    PdfReader = None


# ----------------------------
# Data structures
# ----------------------------

@dataclass
class Rule:
    """
    A conservative, explainable validation rule extracted from the PDF.
    """
    rule_id: str
    source: str  # "pdf"
    confidence: str  # "high" | "medium" | "low"
    kind: str  # "allowed_values" | "conditional_required" | "conditional_na"
    variables: List[str]  # variables involved (with brackets, e.g. "[educ_post_expand]")
    when: Optional[Dict[str, Any]]  # condition in structured-ish form (see builder funcs)
    assert_: Dict[str, Any]  # what must hold if condition is true
    else_: Optional[Dict[str, Any]]  # what must hold if condition is false
    evidence: str  # snippet for explainability

@dataclass
class Issue:
    severity: str  # "error" | "warning"
    tab_name: str
    variable: str
    message: str
    expected: str
    actual: str
    evidence: str

@dataclass
class TabSummary:
    errors: int
    warnings: int
    ok: int

# ----------------------------
# PDF extraction
# ----------------------------

VAR_PATTERN = re.compile(r"\[[A-Za-z0-9_]+\]")

def _read_pdf_text(pdf_path: str) -> str:
    """
    Extract text from a PDF file.

    This is intentionally simple. For MVP v1:
    - We only need enough text to find variable names and a few simple patterns.
    """
    if pdfplumber is not None:
        text_parts: List[str] = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                text_parts.append(txt)
        return "\n".join(text_parts)

    if PdfReader is None:
        raise RuntimeError("No PDF parser available. Install pdfplumber or PyPDF2.")

    reader = PdfReader(pdf_path)
    text_parts = []
    for page in reader.pages:
        try:
            text_parts.append(page.extract_text() or "")
        except Exception:
            text_parts.append("")
    return "\n".join(text_parts)


def extract_pdf_model(pdf_path: str) -> Dict[str, Any]:
    """
    Extract a 'PDF model' from the uploaded codebook PDF.

    Output:
    {
      "variables": set([...]),
      "rules": [Rule as dict, ...],
      "summary": {...}
    }
    """
    text = _read_pdf_text(pdf_path)
    variables = sorted(set(VAR_PATTERN.findall(text)))

    rules: List[Rule] = []
    rules.extend(_extract_yes_no_rules(text))
    rules.extend(_extract_enumeration_rules(text))
    rules.extend(_extract_land_expro_comp_rule(text))
    rules.extend(_extract_restrict_cleav_and_year_rules(text))

    summary = {
        "variables_detected": len(variables),
        "rules_extracted": len(rules),
        "rules_by_confidence": {
            "high": sum(1 for r in rules if r.confidence == "high"),
            "medium": sum(1 for r in rules if r.confidence == "medium"),
            "low": sum(1 for r in rules if r.confidence == "low"),
        }
    }

    return {
        "variables": variables,
        "rules": [asdict(r) for r in rules],
        "summary": summary,
        "raw_text_sample": "\n".join(text.splitlines()[:120])  # helpful for debugging
    }


def _extract_yes_no_rules(text: str) -> List[Rule]:
    """
    Heuristic: if a variable line includes "(Yes = 1, No = 0)" we enforce allowed values {0,1}.
    This is high-confidence because it's explicit.
    """
    rules: List[Rule] = []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        vars_in_line = VAR_PATTERN.findall(line)
        if not vars_in_line:
            continue

        if "Yes = 1" in line and "No = 0" in line:
            for var in vars_in_line:
                evidence = line.strip()
                rules.append(Rule(
                    rule_id=f"YESNO_{var.strip('[]')}",
                    source="pdf",
                    confidence="high",
                    kind="allowed_values",
                    variables=[var],
                    when=None,
                    assert_={"allowed_values": [0, 1]},
                    else_=None,
                    evidence=evidence
                ))
    return rules


def _extract_enumeration_rules(text: str) -> List[Rule]:
    """
    Heuristic: detect enumerated code options like:
      1. No
      2. Informal restrictions
      3. Formal restrictions

    Real PDFs often wrap long list items onto multiple lines. So we:
    - start at a line containing a variable like "[health_pre_types]"
    - scan forward until we hit the next bullet/variable ("• [") or we get too far
    - collect any lines that look like "<number>." at the start

    This is medium-confidence because PDF extraction can still be messy.
    """
    rules: List[Rule] = []
    lines = text.splitlines()

    i = 0
    while i < len(lines):
        line = lines[i]
        vars_in_line = VAR_PATTERN.findall(line)
        if not vars_in_line:
            i += 1
            continue

        var = vars_in_line[0]

        allowed: List[int] = []
        evidence_lines = [line.strip()]

        # Scan forward for up to ~25 lines or until next bullet variable.
        for j in range(i + 1, min(i + 26, len(lines))):
            nxt = lines[j].strip()

            # Stop if we hit the next variable/bullet.
            if nxt.startswith("• ["):
                break

            m = re.match(r"^\s*(\d+)\.\s*", lines[j])
            if m:
                allowed.append(int(m.group(1)))
                evidence_lines.append(nxt)
            else:
                # Wrapped/continuation line: keep as evidence if we are in a list already.
                if allowed:
                    evidence_lines.append(nxt)

        if len(allowed) >= 2:
            rules.append(Rule(
                rule_id=f"ENUM_{var.strip('[]')}",
                source="pdf",
                confidence="medium",
                kind="allowed_values",
                variables=[var],
                when=None,
                assert_={"allowed_values": sorted(set(allowed))},
                else_=None,
                evidence="\n".join([l for l in evidence_lines if l][:10])
            ))

        i += 1

    return rules


def _extract_land_expro_comp_rule(text: str) -> List[Rule]:
    """
    Special-case rule present in the codebook text:
      [land_post_expro_comp] If Yes to any of the above...

    "Any of the above" refers to the preceding expropriation questions:
      [land_post_expro_dev]
      [land_post_expro_redist]
      [land_post_expro_corr]
      [land_post_expro_other]

    This is *medium confidence* because it depends on PDF ordering, but in this codebook
    it's quite stable.
    """
    # Only add this rule if the key variable appears in the PDF.
    if "[land_post_expro_comp]" not in text:
        return []

    evidence_snippet = ""
    for line in text.splitlines():
        if "[land_post_expro_comp]" in line:
            evidence_snippet = line.strip()
            break

    parent_vars = [
        "[land_post_expro_dev]",
        "[land_post_expro_redist]",
        "[land_post_expro_corr]",
        "[land_post_expro_other]",
    ]

    return [Rule(
        rule_id="LAND_EXPRO_COMP_REQUIRED",
        source="pdf",
        confidence="medium",
        kind="conditional_required",
        variables=parent_vars + ["[land_post_expro_comp]"],
        when={
            "any_equals": [{"var": v, "value": 1} for v in parent_vars],
        },
        assert_={
            "var": "[land_post_expro_comp]",
            "required": True,  # must not be blank/NA
        },
        else_={
            "var": "[land_post_expro_comp]",
            "must_be_blank": True,  # should be blank when condition is false
        },
        evidence=evidence_snippet or "If Yes to any of the above. Were landowners compensated..."
    )]


def _extract_restrict_cleav_and_year_rules(text: str) -> List[Rule]:
    """
    Special-case rules:
      [educ_post_restrict_cleav] If restricted (else write NA) ...
      [educ_post_restrict_year] If restricted: What year(s) ...

    In the template, both are in the Education sheet.

    We interpret "restricted" using:
      [educ_post_restrict_group] numeric code:
        1 = no new restrictions
        2/3 = restricted

    This is medium confidence: it's based on the explicit code options.
    """
    rules: List[Rule] = []
    if "[educ_post_restrict_group]" not in text:
        return rules

    def add_rule(child_var: str, rule_id: str, evidence_contains: str) -> None:
        if child_var not in text:
            return
        evidence = ""
        for line in text.splitlines():
            if child_var in line:
                evidence = line.strip()
                break
        rules.append(Rule(
            rule_id=rule_id,
            source="pdf",
            confidence="medium",
            kind="conditional_na",
            variables=["[educ_post_restrict_group]", child_var],
            when={
                "var": "[educ_post_restrict_group]",
                "in": [2, 3],  # restricted
            },
            assert_={
                "var": child_var,
                "required": True,  # must have a code when restricted
            },
            else_={
                "var": child_var,
                "must_be_na": True,  # else write NA
            },
            evidence=evidence or evidence_contains
        ))

    add_rule("[educ_post_restrict_cleav]", "EDUC_POST_RESTRICT_CLEAV", "If restricted (else write NA)")
    add_rule("[educ_post_restrict_year]", "EDUC_POST_RESTRICT_YEAR", "If restricted")

    return rules


# ----------------------------
# Excel parsing
# ----------------------------

def _normalize_code_value(value: Any) -> Optional[Any]:
    """
    Convert the "Numerical code" cell into a normalized form.

    We accept:
    - int / float numeric codes
    - strings like "1: Yes" -> 1
    - strings like "NA" -> "NA"
    - empty -> None
    """
    if value is None:
        return None

    # Some spreadsheets store numbers as floats (e.g. 1.0). Convert safely.
    if isinstance(value, (int, float)):
        # Treat NaN as empty
        try:
            if value != value:  # NaN check
                return None
        except Exception:
            pass
        return int(value)

    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return None
        if s.upper() == "NA":
            return "NA"
        # "1: Yes" or "1 - Yes"
        m = re.match(r"^\s*(\d+)\s*[:\-]", s)
        if m:
            return int(m.group(1))
        # If it's a pure integer string
        if re.fullmatch(r"\d+", s):
            return int(s)
        return s  # keep as-is (might be text codes)
    return value


def parse_excel_workbook(xlsx_path: str) -> Dict[str, Any]:
    """
    Parse the Excel workbook into a normalized structure.

    Output:
    {
      "path": ...,
      "tabs": {
        "Education": {
            "conflict_title": "Kosovo War (1998 - 1999)",
            "rows": {
              "[educ_pre_restrict_quota]": {
                 "code": 1,
                 "year": 1990,
                 "policy": ...,
                 "narrative": ...   # present but not validated
              }, ...
            }
        }, ...
      }
    }

    We assume the template structure:
      Column B: Variable
      Column C: Numerical code
      Column D: Year(s) implemented
      Column E: Formal policies or laws
      Column F: Narrative
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    tabs: Dict[str, Any] = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        conflict_title = ws["B2"].value  # in sample template and sample Kosovo
        conflict_title = str(conflict_title).strip() if conflict_title is not None else ""

        rows: Dict[str, Dict[str, Any]] = {}

        # Variables start at row 6 in the provided template.
        for r in range(6, ws.max_row + 1):
            var_cell = ws.cell(r, 2).value  # column B
            if var_cell is None:
                continue
            var = str(var_cell).strip()
            if not var.startswith("[") or not var.endswith("]"):
                # We only treat bracketed variable IDs as variables.
                continue

            code = _normalize_code_value(ws.cell(r, 3).value)  # col C
            year = ws.cell(r, 4).value  # col D (we don't over-normalize years)
            policy = ws.cell(r, 5).value  # col E
            narrative = ws.cell(r, 6).value  # col F

            rows[var] = {
                "code": code,
                "year": year,
                "policy": policy,
                "narrative": narrative,
                "row_index": r,
            }

        tabs[sheet_name] = {
            "conflict_title": conflict_title,
            "rows": rows,
        }

    return {"path": xlsx_path, "tabs": tabs}


# ----------------------------
# Validation engine
# ----------------------------

PLACEHOLDER_CONFLICT = "CONFLICT (Start - End)"

def validate_workbook(pdf_model: Dict[str, Any], workbook: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validate the parsed workbook against:
    - structural checks from PRD
    - high/medium confidence rules extracted from the PDF

    Returns a report dict that can be rendered as HTML or JSON.
    """
    pdf_vars = set(pdf_model.get("variables", []))
    rules = [Rule(**r) for r in pdf_model.get("rules", [])]

    all_tabs_report: List[Dict[str, Any]] = []
    total_errors = 0
    total_warnings = 0

    for tab_name, tab in workbook["tabs"].items():
        issues: List[Issue] = []
        ok_checks = 0

        # 1) Structural: conflict title in B2
        ct = tab.get("conflict_title", "").strip()
        if ct == "" or ct == PLACEHOLDER_CONFLICT:
            issues.append(Issue(
                severity="warning",
                tab_name=tab_name,
                variable="(metadata)",
                message="Missing conflict title in cell B2 (expected a conflict name like 'Kosovo War (1998 - 1999)').",
                expected="A non-placeholder conflict title in B2",
                actual=repr(ct),
                evidence="Template expects the conflict title in cell B2."
            ))
        else:
            ok_checks += 1

        # 2) Structural: unrecognized variables
        for var in tab["rows"].keys():
            if var not in pdf_vars:
                issues.append(Issue(
                    severity="warning",
                    tab_name=tab_name,
                    variable=var,
                    message="Variable exists in Excel but was not found in the PDF codebook text (might be a mismatch/typo).",
                    expected="Variable should appear in codebook PDF as [var_name]",
                    actual=var,
                    evidence="This warning is conservative: it only checks whether the bracketed ID appears anywhere in the PDF text."
                ))
            else:
                ok_checks += 1

        # 3) Apply extracted rules (but only if their variables exist in this tab)
        for rule in rules:
            # If the rule's main variable isn't on this tab, skip it.
            # (Example: Education rules should not be applied to Land tab.)
            if not any(v in tab["rows"] for v in rule.variables):
                continue
            # If all variables in this rule are absent, skip.
            if all(v not in tab["rows"] for v in rule.variables):
                continue

            new_issues, new_ok = _apply_rule(rule, tab_name, tab["rows"])
            issues.extend(new_issues)
            ok_checks += new_ok

        # Summaries
        errors = sum(1 for i in issues if i.severity == "error")
        warnings = sum(1 for i in issues if i.severity == "warning")
        total_errors += errors
        total_warnings += warnings

        all_tabs_report.append({
            "tab_name": tab_name,
            "summary": asdict(TabSummary(errors=errors, warnings=warnings, ok=ok_checks)),
            "issues": [asdict(i) for i in issues],
        })

    return {
        "pdf_summary": pdf_model.get("summary", {}),
        "workbook_summary": {
            "tabs_validated": len(workbook["tabs"]),
            "total_errors": total_errors,
            "total_warnings": total_warnings,
        },
        "tabs": all_tabs_report,
    }


def _apply_rule(rule: Rule, tab_name: str, rows: Dict[str, Dict[str, Any]]) -> Tuple[List[Issue], int]:
    """
    Apply a single rule to a single tab.

    Returns:
      - list of issues (errors/warnings)
      - count of "ok checks" (for a simple success metric)
    """
    issues: List[Issue] = []
    ok = 0

    if rule.kind == "allowed_values":
        var = rule.variables[0]
        if var not in rows:
            return issues, ok

        val = rows[var]["code"]
        if val is None or val == "NA":
            # We allow blank/NA because blanks are used for unfulfilled conditionals.
            # If a variable is truly required, a conditional_required rule should catch it.
            return issues, ok

        allowed = set(rule.assert_.get("allowed_values", []))
        if isinstance(val, int) and val in allowed:
            ok += 1
        else:
            issues.append(Issue(
                severity="error" if rule.confidence in ("high", "medium") else "warning",
                tab_name=tab_name,
                variable=var,
                message="Invalid numerical code (not in allowed values from codebook).",
                expected=f"One of {sorted(allowed)}",
                actual=repr(val),
                evidence=rule.evidence
            ))
        return issues, ok

    if rule.kind in ("conditional_required", "conditional_na"):
        # Evaluate condition
        cond_true, cond_unknown = _eval_condition(rule.when, rows)

        target = rule.assert_.get("var")
        if not target or target not in rows:
            return issues, ok

        actual = rows[target]["code"]

        # If we can't evaluate the condition (missing parent values), we only warn
        # when something looks suspicious.
        if cond_unknown:
            # If child has a value, warn that we couldn't verify.
            if actual not in (None, "NA"):
                issues.append(Issue(
                    severity="warning",
                    tab_name=tab_name,
                    variable=target,
                    message="Child variable is filled but parent condition could not be evaluated (missing/blank parent codes).",
                    expected="Either fill parent condition variables or leave this blank/NA until the condition is known.",
                    actual=repr(actual),
                    evidence=rule.evidence
                ))
            return issues, ok

        if cond_true:
            # Condition true: enforce rule.assert_
            if rule.assert_.get("required", False):
                if actual is None or actual == "NA":
                    issues.append(Issue(
                        severity="error" if rule.confidence in ("high", "medium") else "warning",
                        tab_name=tab_name,
                        variable=target,
                        message="Required field is missing because its parent condition is true.",
                        expected="A non-blank numerical code (not NA)",
                        actual=repr(actual),
                        evidence=rule.evidence
                    ))
                else:
                    ok += 1
        else:
            # Condition false: enforce rule.else_
            if rule.else_:
                else_var = rule.else_.get("var", target)
                else_actual = rows[else_var]["code"] if else_var in rows else None

                if rule.else_.get("must_be_blank", False):
                    if else_actual is None:
                        ok += 1
                    elif else_actual == "NA":
                        issues.append(Issue(
                            severity="warning",
                            tab_name=tab_name,
                            variable=else_var,
                            message="Field is NA, but codebook flow prefers leaving it blank when the condition is false.",
                            expected="Blank",
                            actual="NA",
                            evidence=rule.evidence
                        ))
                    else:
                        issues.append(Issue(
                            severity="error" if rule.confidence in ("high", "medium") else "warning",
                            tab_name=tab_name,
                            variable=else_var,
                            message="Field should be blank because parent condition is false, but it is filled.",
                            expected="Blank",
                            actual=repr(else_actual),
                            evidence=rule.evidence
                        ))

                if rule.else_.get("must_be_na", False):
                    if else_actual == "NA":
                        ok += 1
                    elif else_actual is None:
                        # PRD: warn if NA required but blank used
                        issues.append(Issue(
                            severity="warning",
                            tab_name=tab_name,
                            variable=else_var,
                            message="Codebook says to write NA when condition is false, but the cell is blank.",
                            expected="NA",
                            actual="Blank",
                            evidence=rule.evidence
                        ))
                    else:
                        issues.append(Issue(
                            severity="error" if rule.confidence in ("high", "medium") else "warning",
                            tab_name=tab_name,
                            variable=else_var,
                            message="Codebook says to write NA when condition is false, but the cell is filled with another value.",
                            expected="NA",
                            actual=repr(else_actual),
                            evidence=rule.evidence
                        ))
        return issues, ok

    # Unknown rule kind: ignore in MVP
    return issues, ok


def _eval_condition(when: Optional[Dict[str, Any]], rows: Dict[str, Dict[str, Any]]) -> Tuple[bool, bool]:
    """
    Evaluate a very small "condition language".

    Returns (condition_is_true, condition_is_unknown)

    Supported forms:
    - {"any_equals": [{"var": "[x]", "value": 1}, ...]}
    - {"var": "[x]", "in": [2,3]}
    """
    if not when:
        return True, False

    # any_equals
    if "any_equals" in when:
        any_list = when["any_equals"]
        unknown = False
        for item in any_list:
            var = item["var"]
            want = item["value"]
            if var not in rows:
                unknown = True
                continue
            val = rows[var]["code"]
            if val is None:
                unknown = True
                continue
            if val == want:
                return True, False
        return False, unknown

    # var in [..]
    if "var" in when and "in" in when:
        var = when["var"]
        allowed = set(when["in"])
        if var not in rows:
            return False, True
        val = rows[var]["code"]
        if val is None:
            return False, True
        if val == "NA":
            return False, True
        if isinstance(val, int) and val in allowed:
            return True, False
        return False, False

    # Unknown condition type -> treat as unknown to be conservative
    return False, True