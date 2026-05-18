"""
PSBER_MFRP mapper — Poland State Budget Execution Report

PDF layout (1 page, 1 table, 4 columns):
  Col 1: Specification (row labels, left-aligned, x0 < 370)
  Col 2: Plan for YYYY (values in mln PLN, European format: "647 200,0")
  Col 3: Execution (same format, can be negative: "-89 294,3")
  Col 4: 3:2 ratio in % (e.g. "28,9"); dash "-" = NA

Each monthly PDF covers "January - [Month] [Year]"; period = YYYY-MM of that last month.

Multi-line labels: some labels wrap across 2-3 y-rows in the PDF. The data values
appear on the last y-row of the label group. This parser accumulates pending label
words from rows with no data and flushes them when a data row arrives.

Continuation noise: "+6.1.7-6.1.8-6.1.9)" follows the "6.1 Domestic" row as a
trailing formula line — it starts with "+" and is skipped.

114 series + period column → loaded from headers.json.
"""

import json
import os
import re
from collections import defaultdict

import pandas as pd
import pdfplumber

if hasattr(__import__("sys").stdout, "reconfigure"):
    __import__("sys").stdout.reconfigure(encoding="utf-8", errors="replace")

with open(os.path.join(os.path.dirname(__file__), "headers.json"), encoding="utf-8") as _f:
    COLUMN_HEADERS = json.load(_f)

# Column x-boundary thresholds (points)
_PLAN_MIN  = 370
_EXEC_MIN  = 455
_RATIO_MIN = 530
_DATA_Y    = 215   # y below which data rows start

_MONTH_MAP = {
    "january": "01", "february": "02", "march": "03", "april": "04",
    "may": "05", "june": "06", "july": "07", "august": "08",
    "september": "09", "october": "10", "november": "11", "december": "12",
}


# ── Period extraction ─────────────────────────────────────────────────────────

def _extract_period(text: str) -> str | None:
    """
    Extract YYYY-MM from the PDF title line:
      "Estimated execution of the state budget in January-April 2026"
    Returns '2026-04'.  Handles single-month "in January 2026" → '2026-01'.
    """
    line_m = re.search(r"Estimated execution of the state budget in[^\n]+", text, re.IGNORECASE)
    if not line_m:
        return None
    line = line_m.group()
    months = re.findall(
        r"(January|February|March|April|May|June|July|August|September"
        r"|October|November|December)",
        line, re.IGNORECASE
    )
    year_m = re.search(r"(\d{4})", line)
    if months and year_m:
        return f"{year_m.group(1)}-{_MONTH_MAP[months[-1].lower()]}"
    return None


# ── Label → series code suffix ────────────────────────────────────────────────

def _match_label(raw: str) -> str | None:
    """
    Map a PDF row label to the shared suffix used in all three prefix groups
    (PLAN / EXECUTION / EXECUTIONRATIO).
    Returns e.g. 'REVENUE.M' or None for skip rows.
    """
    s = raw.lower().strip()
    # Strip leading section numbers / letter bullets / dashes
    s = re.sub(r"^\d+[a-z]?\.\d+\.\d+\s*", "", s)   # 6.1.1
    s = re.sub(r"^\d+[a-z]?\.\d+\s*",       "", s)   # 6.1
    s = re.sub(r"^\d+[a-z]?\.\s*",           "", s)   # 1.  4a.
    s = re.sub(r"^[a-z]\)\s*",               "", s)   # a)  b)
    s = re.sub(r"^-\s*",                     "", s)   # leading dash
    s = re.sub(r"^in\s+which:?\s*-?\s*",     "", s)   # "in which: -"
    s = s.strip()

    if not s or "which:" in s:
        return None

    # More-specific patterns first
    if "refund for the social insurance" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.REFUNDSOCIALINSURANCEFUND.M"
    if "open pension fund revenues" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.OPFREVENUESFROMPRIVATIZATION.M"
    if "distribution of revenues from privatization" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.DISTRIBREVENUESFROMPRIVATIZATION.M"
    if "revenues from privatization" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.REVENUESFROMPRIVATIZATION.M"
    if "pre-financing tasks" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.PREFINTASKSEUFUNDS.M"
    if "central current account" in s or "funds on the central" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.FUNDSONTHECENTRCURACCOUNTBUDGT.M"
    if "management of public finance" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.MANAGEMENTOFPUBLFINLIQ.M"
    if "granted loans" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.GRANTEDLOANS.M"
    if "other revenues and outlays" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.OTHERREVANDOUTLAYS.M"
    if "financial resources transfered from" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.FINRESFROMPREVYEAR.M"
    if "treasury bills" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.TREASURYBILLS.M"
    if "treasury bonds" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.TREASURYBONDS.M"
    if "deposit" in s:
        return "SOURCESFINDEFICIT.DOMESTICDEPOSIT.M"
    if "domestic" in s:
        return "SOURCESFINDEFICIT.DOMESTIC.M"
    if "sources of financing deficit" in s:
        return "SOURCESFINDEFICIT.M"
    if "financial result of the budget of european" in s:
        return "FINRESULTSOFTHEBUDGETEUFUNDS.M"
    if "return of funds transfered" in s:
        return "RETURNFINRESTRASFFORFINEUFUNDS.M"
    if "financial resources transfered for financing" in s:
        return "FINRESTRASFFORFINEUFUNDS.M"
    if "deficit" in s and "surplus" in s:
        return "DEFICITSURPLUS.M"
    if "general subsidies" in s:
        return "EXPENDITURE.GENERALSUBSUDIESFORLOCALGOVUNITS.M"
    if "social insurance fund" in s:
        return "EXPENDITURE.TRASFERSTOSOCIALINSURANCEFUND.M"
    if "pension and disability fund" in s:
        return "EXPENDITURE.TRASFERSTOPENSIONANDDISABILITYFUND.M"
    if "contribution to the european" in s:
        return "EXPENDITURE.CONTRIBUTIONTOEU.M"
    if "debt servicing" in s or ("treasury" in s and "debt" in s):
        return "EXPENDITURE.STATESTREASURYDEBT.M"
    if s == "expenditure" or s.startswith("expenditure "):
        return "EXPENDITURE.M"
    if "structural funds" in s:
        return "REVENUE.NONREFUNDABLE.STRUCTURALFUNDSANDOTHERS.M"
    if "common agriculture policy" in s or "fishery policy" in s:
        return "REVENUE.NONREFUNDABLE.AGRICULTUREFISHERYPOLICY.M"
    if "non-refundable" in s:
        return "REVENUE.NONREFUNDABLE.M"
    if "custom duty" in s:
        return "REVENUE.NONTAXPROCEEDSOFCUSTOMDUTY.M"
    if "non-tax revenue" in s:
        return "REVENUE.NONTAX.M"
    if "financial institutions" in s:
        return "REVENUE.TAX.FININSTITUTIONS.M"
    if "personal income tax" in s:
        return "REVENUE.TAX.PERSONALINCOME.M"
    if "corporate income tax" in s:
        return "REVENUE.TAX.CORPORATEINCOME.M"
    if "excise tax" in s:
        return "REVENUE.TAX.INDIRECTEXCISE.M"
    if "indirect taxes" in s:
        return "REVENUE.TAX.INDIRECT.M"
    if "tax revenue" in s:
        return "REVENUE.TAX.M"
    if "foreign" in s and "european" not in s:
        return "SOURCESFINDEFICIT.FOREIGN.M"
    if s == "revenue":
        return "REVENUE.M"
    return None


# ── Value parsing ─────────────────────────────────────────────────────────────

def _parse_value(raw: str) -> "float | str | None":
    """
    Parse European-format number '647 200,0' → 647200.0.
    Dash '-' means source explicitly says Not Available → returns 'NA' string
    (matching the reference file convention).
    Empty string → None (series absent from this PDF).
    """
    s = raw.strip()
    if not s:
        return None
    if s == "-":
        return "NA"
    s = s.replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# ── PDF parsing ───────────────────────────────────────────────────────────────

def _parse_pdf(path: str) -> tuple[str | None, dict]:
    """
    Parse one PDF file.
    Returns (period_str, data_dict) where data_dict maps full PSBER codes to float|None.
    """
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        words = page.extract_words(x_tolerance=3, y_tolerance=3)

    period = _extract_period(text)

    # Group words into visual rows: sort by y, then cluster words whose y is
    # within 3 px of the first word in the current cluster.
    # This handles the ~0.8 px offset between label baselines and data baselines.
    Y_TOL = 3
    data_words = sorted([w for w in words if w["top"] >= _DATA_Y], key=lambda w: w["top"])

    clusters: list[list] = []
    current: list = []
    for w in data_words:
        if not current or w["top"] - current[0]["top"] <= Y_TOL:
            current.append(w)
        else:
            clusters.append(current)
            current = [w]
    if current:
        clusters.append(current)

    rows: list[dict] = []
    for cluster in clusters:
        row: dict = {"label": [], "plan": [], "exec": [], "ratio": []}
        for w in cluster:
            if w["x0"] < _PLAN_MIN:
                row["label"].append(w["text"])
            elif w["x0"] < _EXEC_MIN:
                row["plan"].append(w["text"])
            elif w["x0"] < _RATIO_MIN:
                row["exec"].append(w["text"])
            else:
                row["ratio"].append(w["text"])
        rows.append(row)

    data: dict[str, float | None] = {}
    pending_label: list[str] = []

    for row in rows:
        has_data = bool(row["plan"] or row["exec"] or row["ratio"])
        label_text = " ".join(row["label"]).strip()

        if not has_data:
            # Skip "of which:", "in which:", formula continuations starting with "+"
            if (not label_text
                    or "which:" in label_text.lower()
                    or label_text.startswith("+")):
                continue
            pending_label.append(label_text)
            continue

        # Data row: combine any accumulated label with current row's label
        full_label = " ".join(pending_label + ([label_text] if label_text else [])).strip()
        pending_label = []

        suffix = _match_label(full_label)
        if suffix is None:
            continue

        plan_val  = _parse_value(" ".join(row["plan"]))
        exec_val  = _parse_value(" ".join(row["exec"]))
        ratio_val = _parse_value(" ".join(row["ratio"]))

        data[f"PSBER.PLAN.{suffix}"]          = plan_val
        data[f"PSBER.EXECUTION.{suffix}"]     = exec_val
        data[f"PSBER.EXECUTIONRATIO.{suffix}"] = ratio_val

    matched = sum(1 for k in data if isinstance(data[k], (int, float)))
    print(f"[mapper] {os.path.basename(path)} → period={period}, "
          f"series={len(data)//3}, non-null values={matched}")
    return period, data


# ── Output row assembly ───────────────────────────────────────────────────────

def _record_to_row(period: str, data: dict) -> list:
    codes = COLUMN_HEADERS["codes"]
    row = [period] + [None] * (len(codes) - 1)
    for i, code in enumerate(codes[1:], start=1):
        val = data.get(code)
        if val is None:
            continue
        if val == "NA":
            row[i] = "NA"
        else:
            try:
                row[i] = float(val)
            except (ValueError, TypeError):
                row[i] = val
    return row


# ── Public interface ──────────────────────────────────────────────────────────

def map_to_output(pdf_paths: list[tuple[str | None, str]],
                  existing_path: str | None = None) -> pd.DataFrame:
    """
    Parse PDF files and merge into existing history.

    pdf_paths:     list of (period_hint, local_pdf_path)
    existing_path: path to existing DATA xlsx (rows 0/1 = headers, row 2+ = data)

    Returns DataFrame with 2 header rows + date-sorted data rows.
    """
    existing_rows: dict[str, list] = {}
    if existing_path and os.path.exists(existing_path):
        ex = pd.read_excel(existing_path, header=None)
        for _, r in ex.iloc[2:].iterrows():
            period = str(r.iloc[0]) if pd.notna(r.iloc[0]) else None
            if period:
                existing_rows[period] = list(r)
        print(f"[mapper] Loaded {len(existing_rows)} existing rows from {existing_path}")

    for _hint, pdf_path in pdf_paths:
        period, data = _parse_pdf(pdf_path)
        if period is None:
            print(f"[mapper] WARNING: could not determine period for {pdf_path}, skipping")
            continue
        new_row = _record_to_row(period, data)
        if period in existing_rows:
            merged = list(existing_rows[period])
            for i, val in enumerate(new_row):
                if val is not None:
                    merged[i] = val
            existing_rows[period] = merged
        else:
            existing_rows[period] = new_row

    sorted_rows = [existing_rows[d] for d in sorted(existing_rows)]
    all_rows = [COLUMN_HEADERS["codes"], COLUMN_HEADERS["descriptions"]] + sorted_rows
    return pd.DataFrame(all_rows)


def build_metadata_rows() -> list[dict]:
    """Return metadata dicts for each series — used by main.py for META xlsx."""
    codes = COLUMN_HEADERS["codes"][1:]
    descs = COLUMN_HEADERS["descriptions"][1:]
    rows = []
    for code, desc in zip(codes, descs):
        unit = "%" if "EXECUTIONRATIO" in code else "mln PLN"
        rows.append({
            "CODE":              code,
            "DESCRIPTION":       desc,
            "FREQUENCY":         "Monthly",
            "UNIT":              unit,
            "SOURCE_NAME":       "Ministry of Finance Republic of Poland",
            "SOURCE_URL":        PAGE_URL,
            "DATASET":           "PSBER_MFRP",
            "NEXT_RELEASE_DATE": "",
        })
    return rows


PAGE_URL = "https://www.gov.pl/web/finance/estimated-execution-of-the-state-budget"
