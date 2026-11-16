#!/usr/bin/env python3
"""
Build standardized historical financial dataset + Excel model template from a CSV.

Input: a CSV containing Income Statement, Balance Sheet, Cash Flow, and KPIs.
Output:
  1) cleaned_standardized.csv (long-form standardized dataset)
  2) model_inputs.xlsx (tabs formatted to match an IB-style model template)

Validation included:
  - Totals / statement checks (IS, BS, CF tie-outs where possible)
  - YoY growth rates
  - Missing value detection
  - Outlier flags (robust MAD z-score)

Usage:
  python build_model_inputs.py --in historicals.csv --outdir ./out

CSV formats supported:
  A) Long format (preferred):
     Statement, LineItem, Year, Value [, Entity, Currency, Units]
  B) Wide format:
     Statement, LineItem, FY2019, FY2020, FY2021, ...
     (or 2019,2020,...)

Notes:
  - This script standardizes line item names via a mapping dictionary you can extend.
  - It does NOT invent numbers; it cleans/standardizes whatever is in the CSV.
"""

from __future__ import annotations

import argparse
import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# -----------------------------
# Configuration / Standards
# -----------------------------

# Canonical statements
STMT_IS = "Income Statement"
STMT_BS = "Balance Sheet"
STMT_CF = "Cash Flow"
STMT_KPI = "KPI"

# Canonical line items (you can expand)
CANON = {
    # Income statement
    "Revenue": ("IS", 1),
    "COGS": ("IS", 2),
    "Gross Profit": ("IS", 3),
    "Labor / Store Opex": ("IS", 4),
    "Other Opex / SG&A": ("IS", 5),
    "SBC": ("IS", 6),
    "D&A": ("IS", 7),
    "EBIT": ("IS", 8),
    "Interest Expense": ("IS", 9),
    "EBT": ("IS", 10),
    "Taxes": ("IS", 11),
    "Net Income": ("IS", 12),

    # Balance sheet (simplified)
    "Cash": ("BS", 1),
    "Accounts Receivable": ("BS", 2),
    "Inventory": ("BS", 3),
    "Other Current Assets": ("BS", 4),
    "Total Current Assets": ("BS", 5),
    "PP&E Net": ("BS", 6),
    "Other Long-Term Assets": ("BS", 7),
    "Total Assets": ("BS", 8),

    "Accounts Payable": ("BS", 9),
    "Accrued / Other Current Liabilities": ("BS", 10),
    "Total Current Liabilities": ("BS", 11),
    "Long-Term Debt": ("BS", 12),
    "Other Long-Term Liabilities": ("BS", 13),
    "Total Liabilities": ("BS", 14),

    "Common Equity": ("BS", 15),
    "Retained Earnings": ("BS", 16),
    "Total Equity": ("BS", 17),
    "Total Liabilities & Equity": ("BS", 18),

    # Cash flow (simplified)
    "Net Income (CF)": ("CF", 1),
    "D&A (CF)": ("CF", 2),
    "ΔNWC": ("CF", 3),
    "Capex": ("CF", 4),
    "Unlevered FCF": ("CF", 5),
    "Interest (CF)": ("CF", 6),
    "Debt Issued/(Repaid)": ("CF", 7),
    "ΔCash": ("CF", 8),
    "Beginning Cash": ("CF", 9),
    "Ending Cash": ("CF", 10),
}

# Common aliases -> canonical line items
# Extend these based on your CSV exports.
ALIASES: Dict[str, str] = {
    # Revenue
    "total revenue": "Revenue",
    "net revenue": "Revenue",
    "revenues": "Revenue",
    "sales": "Revenue",

    # COGS
    "cost of goods sold": "COGS",
    "cost of sales": "COGS",

    # Gross profit
    "gross profit": "Gross Profit",

    # Opex
    "labor": "Labor / Store Opex",
    "store labor": "Labor / Store Opex",
    "store operating expenses": "Labor / Store Opex",
    "operating expenses": "Other Opex / SG&A",
    "sg&a": "Other Opex / SG&A",
    "selling general and administrative": "Other Opex / SG&A",

    # SBC
    "stock-based compensation": "SBC",
    "share-based compensation": "SBC",

    # D&A
    "depreciation and amortization": "D&A",
    "depreciation & amortization": "D&A",
    "da": "D&A",
    "d&a": "D&A",

    # EBIT
    "operating income": "EBIT",
    "ebit": "EBIT",

    # Interest
    "interest expense": "Interest Expense",
    "net interest expense": "Interest Expense",

    # Taxes
    "income tax": "Taxes",
    "income taxes": "Taxes",
    "taxes": "Taxes",

    # NI
    "net income": "Net Income",
    "net earnings": "Net Income",

    # Balance sheet
    "cash and cash equivalents": "Cash",
    "accounts receivable": "Accounts Receivable",
    "inventories": "Inventory",
    "inventory": "Inventory",
    "other current assets": "Other Current Assets",
    "total current assets": "Total Current Assets",
    "property plant and equipment": "PP&E Net",
    "pp&e": "PP&E Net",
    "other assets": "Other Long-Term Assets",
    "total assets": "Total Assets",
    "accounts payable": "Accounts Payable",
    "accrued liabilities": "Accrued / Other Current Liabilities",
    "other current liabilities": "Accrued / Other Current Liabilities",
    "total current liabilities": "Total Current Liabilities",
    "long-term debt": "Long-Term Debt",
    "total liabilities": "Total Liabilities",
    "common stock": "Common Equity",
    "retained earnings": "Retained Earnings",
    "total equity": "Total Equity",
    "total liabilities and equity": "Total Liabilities & Equity",

    # Cash flow
    "net income (loss)": "Net Income (CF)",
    "net income - cash flow": "Net Income (CF)",
    "depreciation and amortization - cash flow": "D&A (CF)",
    "change in working capital": "ΔNWC",
    "changes in working capital": "ΔNWC",
    "capital expenditures": "Capex",
    "capex": "Capex",
    "free cash flow": "Unlevered FCF",
    "unlevered free cash flow": "Unlevered FCF",
    "net change in cash": "ΔCash",
    "beginning cash": "Beginning Cash",
    "ending cash": "Ending Cash",
}

# Statement normalization
STMT_ALIASES = {
    "is": STMT_IS,
    "income statement": STMT_IS,
    "income": STMT_IS,
    "p&l": STMT_IS,
    "pl": STMT_IS,

    "bs": STMT_BS,
    "balance sheet": STMT_BS,
    "balance": STMT_BS,

    "cf": STMT_CF,
    "cash flow": STMT_CF,
    "cashflow": STMT_CF,
    "statement of cash flows": STMT_CF,

    "kpi": STMT_KPI,
    "kpis": STMT_KPI,
    "metrics": STMT_KPI,
}

# Excel template row layouts (line item ordering)
TEMPLATE_IS = [
    "Revenue",
    "COGS",
    "Gross Profit",
    "Labor / Store Opex",
    "Other Opex / SG&A",
    "SBC",
    "D&A",
    "EBIT",
    "Interest Expense",
    "EBT",
    "Taxes",
    "Net Income",
]

TEMPLATE_BS = [
    "Cash",
    "Accounts Receivable",
    "Inventory",
    "Other Current Assets",
    "Total Current Assets",
    "PP&E Net",
    "Other Long-Term Assets",
    "Total Assets",
    "Accounts Payable",
    "Accrued / Other Current Liabilities",
    "Total Current Liabilities",
    "Long-Term Debt",
    "Other Long-Term Liabilities",
    "Total Liabilities",
    "Common Equity",
    "Retained Earnings",
    "Total Equity",
    "Total Liabilities & Equity",
]

TEMPLATE_CF = [
    "Net Income (CF)",
    "D&A (CF)",
    "ΔNWC",
    "Capex",
    "Unlevered FCF",
    "Interest (CF)",
    "Debt Issued/(Repaid)",
    "ΔCash",
    "Beginning Cash",
    "Ending Cash",
]

# -----------------------------
# Helpers
# -----------------------------

def _clean_text(x: str) -> str:
    x = "" if x is None else str(x)
    x = x.strip().lower()
    x = re.sub(r"\s+", " ", x)
    return x

def normalize_statement(stmt_raw: str) -> str:
    key = _clean_text(stmt_raw)
    return STMT_ALIASES.get(key, stmt_raw.strip() if isinstance(stmt_raw, str) else str(stmt_raw))

def normalize_line_item(item_raw: str) -> str:
    key = _clean_text(item_raw)
    # Strip punctuation often found in exports
    key2 = re.sub(r"[^a-z0-9&/\-\s]", "", key)
    key2 = re.sub(r"\s+", " ", key2).strip()
    return ALIASES.get(key2, item_raw.strip() if isinstance(item_raw, str) else str(item_raw))

def detect_year_columns(columns: List[str]) -> List[str]:
    year_cols = []
    for c in columns:
        s = str(c).strip()
        # Accept FY2021, 2021, 2021A, etc.
        m = re.search(r"(?:FY)?(\d{4})", s)
        if m:
            year_cols.append(c)
    return year_cols

def parse_year_value_from_colname(colname: str) -> Optional[int]:
    s = str(colname).strip()
    m = re.search(r"(?:FY)?(\d{4})", s)
    if not m:
        return None
    return int(m.group(1))

def robust_mad_zscore(series: pd.Series) -> pd.Series:
    """Robust z-score using median absolute deviation."""
    s = pd.to_numeric(series, errors="coerce")
    med = np.nanmedian(s)
    mad = np.nanmedian(np.abs(s - med))
    if mad == 0 or np.isnan(mad):
        return pd.Series([np.nan] * len(s), index=s.index)
    return 0.6745 * (s - med) / mad

# -----------------------------
# Parsing + Standardization
# -----------------------------

def read_input_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]
    return df

def to_long_form(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(df.columns)
    lower_cols = [_clean_text(c) for c in cols]

    # Case A: already long
    required_long = {"statement", "lineitem", "year", "value"}
    if required_long.issubset(set(lower_cols)):
        # Map original column names to canonical
        col_map = {cols[lower_cols.index("statement")]: "Statement",
                   cols[lower_cols.index("lineitem")]: "LineItem",
                   cols[lower_cols.index("year")]: "Year",
                   cols[lower_cols.index("value")]: "Value"}
        out = df.rename(columns=col_map).copy()
        return out[["Statement", "LineItem", "Year", "Value"] + [c for c in out.columns if c not in col_map.values()]]

    # Case B: wide years
    if "statement" in lower_cols and "lineitem" in lower_cols:
        stmt_col = cols[lower_cols.index("statement")]
        li_col = cols[lower_cols.index("lineitem")]
        year_cols = detect_year_columns(cols)
        if not year_cols:
            raise ValueError("Could not detect Year columns. Provide long format with Statement,LineItem,Year,Value.")
        id_vars = [stmt_col, li_col] + [c for c in cols if c not in year_cols and c not in [stmt_col, li_col]]
        melted = df.melt(id_vars=id_vars, value_vars=year_cols, var_name="Year", value_name="Value")
        melted = melted.rename(columns={stmt_col: "Statement", li_col: "LineItem"})
        # Convert Year column name to numeric year where possible
        melted["Year"] = melted["Year"].apply(parse_year_value_from_colname)
        return melted

    raise ValueError("CSV must include either long format columns (Statement,LineItem,Year,Value) or wide format with Statement,LineItem and year columns.")

def standardize_dataset(df_long: pd.DataFrame) -> pd.DataFrame:
    out = df_long.copy()

    # Clean and normalize
    out["Statement"] = out["Statement"].apply(normalize_statement)
    out["LineItem_raw"] = out["LineItem"].astype(str)
    out["LineItem"] = out["LineItem"].apply(normalize_line_item)

    out["Year"] = pd.to_numeric(out["Year"], errors="coerce").astype("Int64")
    out["Value"] = pd.to_numeric(out["Value"], errors="coerce")

    # Drop rows without a valid year
    out = out.dropna(subset=["Year"]).copy()

    # Basic unit standardization placeholder
    # If your CSV includes Units/Currency columns, you can normalize here.

    # Add canonical statement code (IS/BS/CF/KPI) for sorting
    stmt_code = {
        STMT_IS: "IS",
        STMT_BS: "BS",
        STMT_CF: "CF",
        STMT_KPI: "KPI",
    }
    out["StmtCode"] = out["Statement"].map(stmt_code).fillna("OTHER")

    # Sort helper rank where known
    def rank_row(r: pd.Series) -> int:
        canon = CANON.get(r["LineItem"])
        if canon:
            return canon[1]
        return 999

    out["LineRank"] = out.apply(rank_row, axis=1)

    # Deduplicate: if duplicates exist (same stmt, line, year) take last non-null
    out = out.sort_values(["Statement", "LineItem", "Year"]).copy()
    out["Value"] = out.groupby(["Statement", "LineItem", "Year"])["Value"].transform(lambda s: s.ffill().bfill())
    out = out.drop_duplicates(subset=["Statement", "LineItem", "Year"], keep="last")

    # Ensure sign conventions (optional):
    # Many exports have COGS/Capex as positive numbers; IB models often show expenses as positive too.
    # We'll keep as-is. If you want enforced convention, add rules here.

    # Add YoY growth for each line item within statement
    out = out.sort_values(["Statement", "LineItem", "Year"])
    out["YoY_Growth"] = out.groupby(["Statement", "LineItem"])["Value"].pct_change()

    # Missing flag
    out["Missing_Flag"] = out["Value"].isna()

    # Outlier flag (robust) by line item
    out["Outlier_Z"] = out.groupby(["Statement", "LineItem"])["Value"].transform(robust_mad_zscore)
    out["Outlier_Flag"] = out["Outlier_Z"].abs() >= 4  # robust threshold

    return out

# -----------------------------
# Validation / Checks
# -----------------------------

@dataclass
class ValidationResult:
    checks: pd.DataFrame
    issues: pd.DataFrame

def build_validation(standard: pd.DataFrame) -> ValidationResult:
    """
    Produces:
      - checks: per-year statement checks where possible
      - issues: row-level issues (missing, outlier, duplicates already handled)
    """
    # Row-level issues
    issues = standard.loc[standard["Missing_Flag"] | standard["Outlier_Flag"],
                          ["Statement", "LineItem", "Year", "Value", "Missing_Flag", "Outlier_Flag", "Outlier_Z", "LineItem_raw"]].copy()

    # Pivot for checks
    pivot = standard.pivot_table(index="Year", columns=["Statement", "LineItem"], values="Value", aggfunc="last")

    rows = []
    years = sorted([y for y in pivot.index if pd.notna(y)])

    def get(stmt: str, item: str, y: int) -> Optional[float]:
        try:
            return float(pivot.loc[y, (stmt, item)])
        except Exception:
            return None

    for y in years:
        # Income Statement checks (if components exist)
        rev = get(STMT_IS, "Revenue", y)
        cogs = get(STMT_IS, "COGS", y)
        gp = get(STMT_IS, "Gross Profit", y)
        labor = get(STMT_IS, "Labor / Store Opex", y)
        opex = get(STMT_IS, "Other Opex / SG&A", y)
        sbc = get(STMT_IS, "SBC", y)
        da = get(STMT_IS, "D&A", y)
        ebit = get(STMT_IS, "EBIT", y)
        ie = get(STMT_IS, "Interest Expense", y)
        ebt = get(STMT_IS, "EBT", y)
        tax = get(STMT_IS, "Taxes", y)
        ni = get(STMT_IS, "Net Income", y)

        # Treat checks as "difference" = should be ~0 (tolerance applied in Excel later)
        if rev is not None and cogs is not None and gp is not None:
            rows.append({"Year": y, "Check": "IS: Gross Profit = Revenue - COGS",
                         "Diff": gp - (rev - cogs)})

        if all(v is not None for v in [gp, labor, opex, sbc, da, ebit]):
            rows.append({"Year": y, "Check": "IS: EBIT = GP - Labor - OtherOpex - SBC - D&A",
                         "Diff": ebit - (gp - labor - opex - sbc - da)})

        if all(v is not None for v in [ebit, ie, ebt]):
            rows.append({"Year": y, "Check": "IS: EBT = EBIT - Interest",
                         "Diff": ebt - (ebit - ie)})

        if all(v is not None for v in [ebt, tax, ni]):
            rows.append({"Year": y, "Check": "IS: Net Income = EBT - Taxes",
                         "Diff": ni - (ebt - tax)})

        # Balance Sheet check
        ta = get(STMT_BS, "Total Assets", y)
        tle = get(STMT_BS, "Total Liabilities & Equity", y)
        if ta is not None and tle is not None:
            rows.append({"Year": y, "Check": "BS: Total Assets = Total Liabilities & Equity",
                         "Diff": ta - tle})

        # Cash Flow check: Ending cash tie (Ending = Beginning + ΔCash)
        beg = get(STMT_CF, "Beginning Cash", y)
        end = get(STMT_CF, "Ending Cash", y)
        dch = get(STMT_CF, "ΔCash", y)
        if all(v is not None for v in [beg, end, dch]):
            rows.append({"Year": y, "Check": "CF: Ending Cash = Beginning Cash + ΔCash",
                         "Diff": end - (beg + dch)})

        # Optional: UFCF bridge if inputs exist (NOPAT not in dataset; so minimal check)
        ni_cf = get(STMT_CF, "Net Income (CF)", y)
        da_cf = get(STMT_CF, "D&A (CF)", y)
        dnwc = get(STMT_CF, "ΔNWC", y)
        capex = get(STMT_CF, "Capex", y)
        ufcf = get(STMT_CF, "Unlevered FCF", y)
        if all(v is not None for v in [ni_cf, da_cf, dnwc, capex, ufcf]):
            # This is not a perfect UFCF definition, but a common simplified bridge:
            # UFCF ≈ NI + D&A - ΔNWC - Capex (ignoring interest/tax normalization)
            rows.append({"Year": y, "Check": "CF: UFCF approx = NI + D&A - ΔNWC - Capex (simplified)",
                         "Diff": ufcf - (ni_cf + da_cf - dnwc - capex)})

    checks = pd.DataFrame(rows)
    if checks.empty:
        checks = pd.DataFrame(columns=["Year", "Check", "Diff"])

    return ValidationResult(checks=checks, issues=issues)

# -----------------------------
# Excel Writer (Template)
# -----------------------------

def style_workbook(wb: Workbook) -> None:
    # Basic styles
    thin = Side(border_style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill("solid", fgColor="1F4E79")  # dark blue
    header_font = Font(color="FFFFFF", bold=True)
    header_align = Alignment(horizontal="center", vertical="center")

    section_fill = PatternFill("solid", fgColor="F2F2F2")
    section_font = Font(bold=True)

    input_font = Font(color="1F4E79")  # blue
    note_font = Font(color="666666", italic=True)

    wb._style_cache = {
        "border": border,
        "header_fill": header_fill,
        "header_font": header_font,
        "header_align": header_align,
        "section_fill": section_fill,
        "section_font": section_font,
        "input_font": input_font,
        "note_font": note_font,
    }

def set_col_widths(ws, widths: Dict[int, float]) -> None:
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

def write_header_row(ws, row: int, headers: List[str]) -> None:
    sc = ws.parent._style_cache
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=j, value=h)
        cell.fill = sc["header_fill"]
        cell.font = sc["header_font"]
        cell.alignment = sc["header_align"]
        cell.border = sc["border"]
    ws.freeze_panes = ws.cell(row=row + 1, column=1)

def write_section(ws, row: int, title: str) -> int:
    sc = ws.parent._style_cache
    c = ws.cell(row=row, column=1, value=title)
    c.fill = sc["section_fill"]
    c.font = sc["section_font"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    return row + 1

def years_from_standard(standard: pd.DataFrame) -> List[int]:
    yrs = sorted(standard["Year"].dropna().unique().tolist())
    yrs_int = [int(y) for y in yrs]
    return yrs_int

def pivot_statement(standard: pd.DataFrame, statement: str, line_items_order: List[str], years: List[int]) -> pd.DataFrame:
    df = standard.loc[standard["Statement"] == statement, ["LineItem", "Year", "Value"]].copy()
    p = df.pivot_table(index="LineItem", columns="Year", values="Value", aggfunc="last")
    # Ensure order + include missing template rows
    out = pd.DataFrame(index=line_items_order)
    for y in years:
        out[y] = p.get(y)
    out.index.name = "LineItem"
    return out

def write_historical_fs_sheet(wb: Workbook, standard: pd.DataFrame) -> None:
    ws = wb.create_sheet("Historical_FS")
    set_col_widths(ws, {1: 40, 2: 28, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14, 11: 14, 12: 14})

    years = years_from_standard(standard)
    # Put years starting col C
    headers = ["Line Item", "Notes"] + [str(y) for y in years]
    write_header_row(ws, 1, headers)

    r = 3
    r = write_section(ws, r, "INCOME STATEMENT (HIST)")
    is_df = pivot_statement(standard, STMT_IS, TEMPLATE_IS, years)
    r = write_table(ws, r, is_df, notes_stub=True)

    r += 1
    r = write_section(ws, r, "BALANCE SHEET (HIST)")
    bs_df = pivot_statement(standard, STMT_BS, TEMPLATE_BS, years)
    r = write_table(ws, r, bs_df, notes_stub=True)

    r += 1
    r = write_section(ws, r, "CASH FLOW (HIST)")
    cf_df = pivot_statement(standard, STMT_CF, TEMPLATE_CF, years)
    r = write_table(ws, r, cf_df, notes_stub=True)

def write_table(ws, start_row: int, df: pd.DataFrame, notes_stub: bool = True) -> int:
    sc = ws.parent._style_cache
    r = start_row
    for i, (idx, row) in enumerate(df.iterrows()):
        ws.cell(row=r, column=1, value=str(idx)).border = sc["border"]
        if notes_stub:
            ws.cell(row=r, column=2, value="").border = sc["border"]
        for j, col in enumerate(df.columns, start=3):
            v = row[col]
            c = ws.cell(row=r, column=j, value=(None if pd.isna(v) else float(v)))
            c.number_format = "#,##0"
            c.border = sc["border"]
        r += 1
    return r

def write_inputs_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Inputs")
    set_col_widths(ws, {1: 42, 2: 18, 3: 18, 4: 45})
    write_header_row(ws, 1, ["Item", "Value", "Name", "Notes"])

    sc = wb._style_cache
    r = 3

    def put(item, name=None, note=None):
        nonlocal r
        ws.cell(row=r, column=1, value=item)
        vcell = ws.cell(row=r, column=2, value=None)
        vcell.font = sc["input_font"]  # blue inputs
        if name:
            ws.cell(row=r, column=3, value=name)
        if note:
            ncell = ws.cell(row=r, column=4, value=note)
            ncell.font = sc["note_font"]
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = sc["border"]
        r += 1

    def section(title):
        nonlocal r
        ws.cell(row=r, column=1, value=title).fill = sc["section_fill"]
        ws.cell(row=r, column=1).font = sc["section_font"]
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        r += 1

    section("MODEL CONTROL PANEL")
    put("Fiscal year end", "Fiscal_Year_End", "Text label, e.g., 'Sep'")
    put("Projection years", "nYears", "Integer number of forecast years")
    put("Exit year index", "Exit_Index", "e.g., 5 for FY5 exit")
    put("Terminal method toggle (1=Gordon,2=Exit)", "TV_Method_Toggle", "1=Gordon Growth, 2=Exit Multiple")

    r += 1
    section("DCF – MARKET ASSUMPTIONS")
    put("Risk-free rate", "Risk_Free_Rate")
    put("Equity risk premium", "Equity_Risk_Premium")
    put("Beta", "Beta")
    put("Pre-tax cost of debt", "Cost_of_Debt")
    put("Marginal tax rate", "Tax_Rate_Marginal")
    put("Terminal growth", "Terminal_Growth")
    put("Exit EV/EBITDA multiple", "Exit_Multiple")
    put("Target debt weight", "Target_Debt_%")
    put("Target equity weight", "Target_Equity_%")

    r += 1
    section("LBO – TRANSACTION ASSUMPTIONS")
    put("Entry EV/EBITDA multiple", "Entry_Multiple")
    put("Transaction fees (% of EV)", "Transaction_Fee_%")
    put("Financing fees (% of new debt)", "Financing_Fee_%")
    put("OID (% of new debt)", "OID_%")

    r += 1
    section("REVOLVER")
    put("Revolver limit", "Revolver_Limit")
    put("Revolver rate", "Revolver_Rate")
    put("Minimum cash", "Min_Cash")

def write_drivers_sheet(wb: Workbook, standard: pd.DataFrame) -> None:
    ws = wb.create_sheet("Drivers")
    set_col_widths(ws, {1: 42, 2: 18, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14, 11: 14, 12: 14})

    years = years_from_standard(standard)
    # We will lay out historical years to the right as a reference, and leave forecast blank in this template.
    headers = ["Driver", "Units"] + [str(y) for y in years] + ["FY1", "FY2", "FY3", "FY4", "FY5"]
    write_header_row(ws, 1, headers)

    sc = wb._style_cache
    r = 3

    # Section: Revenue build
    ws.cell(row=r, column=1, value="REVENUE BUILD").fill = sc["section_fill"]
    ws.cell(row=r, column=1).font = sc["section_font"]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    # Drivers (left blank—user fills)
    rev_drivers = [
        ("Stores – Begin", "count"),
        ("New stores", "count"),
        ("Closures", "count"),
        ("Stores – End", "count"),
        ("Avg stores", "count"),
        ("Traffic growth", "%"),
        ("Ticket growth", "%"),
        ("Price/mix", "%"),
        ("Comp growth", "%"),
        ("Revenue per avg store", "currency"),
        ("Revenue from comps", "currency"),
        ("Revenue from new stores", "currency"),
        ("Total revenue", "currency"),
    ]
    for name, unit in rev_drivers:
        ws.cell(row=r, column=1, value=name).border = sc["border"]
        ws.cell(row=r, column=2, value=unit).border = sc["border"]
        # Forecast entry cells (FY1..FY5) in blue
        for col in range(3 + len(years), 3 + len(years) + 5):
            c = ws.cell(row=r, column=col, value=None)
            c.font = sc["input_font"]
            c.border = sc["border"]
        # History reference columns (if any) left blank
        for col in range(3, 3 + len(years)):
            ws.cell(row=r, column=col, value=None).border = sc["border"]
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="MARGINS & OPEX (AS % OF REV)").fill = sc["section_fill"]
    ws.cell(row=r, column=1).font = sc["section_font"]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1

    margin_drivers = [
        ("COGS % revenue", "%"),
        ("Labor % revenue", "%"),
        ("Other opex % revenue", "%"),
        ("SBC % revenue", "%"),
        ("D&A % revenue", "%"),
        ("Capex % revenue", "%"),
        ("NWC % revenue", "%"),
        ("Cash tax rate", "%"),
    ]
    for name, unit in margin_drivers:
        ws.cell(row=r, column=1, value=name).border = sc["border"]
        ws.cell(row=r, column=2, value=unit).border = sc["border"]
        # History-derived placeholders: compute from standard where possible
        # We'll attempt: (line item / revenue) for relevant items in history.
        hist_vals = compute_hist_margin_series(standard, name, years)
        for j, y in enumerate(years):
            c = ws.cell(row=r, column=3 + j, value=(None if hist_vals is None else hist_vals.get(y)))
            c.number_format = "0.00%"
            c.border = sc["border"]
        # Forecast in blue
        for col in range(3 + len(years), 3 + len(years) + 5):
            c = ws.cell(row=r, column=col, value=None)
            c.font = sc["input_font"]
            c.border = sc["border"]
        r += 1

def compute_hist_margin_series(standard: pd.DataFrame, driver_name: str, years: List[int]) -> Optional[Dict[int, float]]:
    # Map driver name to numerator line item
    mapping = {
        "COGS % revenue": (STMT_IS, "COGS"),
        "Labor % revenue": (STMT_IS, "Labor / Store Opex"),
        "Other opex % revenue": (STMT_IS, "Other Opex / SG&A"),
        "SBC % revenue": (STMT_IS, "SBC"),
        "D&A % revenue": (STMT_IS, "D&A"),
        "Capex % revenue": (STMT_CF, "Capex"),
        "NWC % revenue": (STMT_BS, None),  # needs NWC construction; skip by default
        "Cash tax rate": (STMT_IS, "Taxes"),  # taxes / EBT is more typical; we'll skip if missing
    }

    if driver_name not in mapping:
        return None

    stmt, item = mapping[driver_name]
    # Revenue series
    rev = standard.loc[(standard["Statement"] == STMT_IS) & (standard["LineItem"] == "Revenue"), ["Year", "Value"]]
    if rev.empty:
        return None
    rev_map = {int(y): float(v) for y, v in zip(rev["Year"], rev["Value"]) if pd.notna(y) and pd.notna(v)}

    if driver_name == "NWC % revenue":
        # If AR/AP/Inventory exist, construct NWC = (AR + Inv + OCA) - (AP + Accrued)
        ar = series_map(standard, STMT_BS, "Accounts Receivable")
        inv = series_map(standard, STMT_BS, "Inventory")
        oca = series_map(standard, STMT_BS, "Other Current Assets")
        ap = series_map(standard, STMT_BS, "Accounts Payable")
        accr = series_map(standard, STMT_BS, "Accrued / Other Current Liabilities")
        if not ar and not inv and not oca and not ap and not accr:
            return None
        out = {}
        for y in years:
            if y not in rev_map or rev_map[y] == 0:
                out[y] = None
                continue
            nwc = (ar.get(y, 0.0) + inv.get(y, 0.0) + oca.get(y, 0.0)) - (ap.get(y, 0.0) + accr.get(y, 0.0))
            out[y] = nwc / rev_map[y]
        return out

    if driver_name == "Cash tax rate":
        # Taxes / EBT (if EBT exists); otherwise skip
        taxes = series_map(standard, STMT_IS, "Taxes")
        ebt = series_map(standard, STMT_IS, "EBT")
        if not taxes or not ebt:
            return None
        out = {}
        for y in years:
            if y not in taxes or y not in ebt or ebt[y] == 0:
                out[y] = None
            else:
                out[y] = taxes[y] / ebt[y]
        return out

    # Standard margin = numerator / revenue
    num = series_map(standard, stmt, item) if item else {}
    if not num:
        return None
    out = {}
    for y in years:
        if y not in rev_map or rev_map[y] == 0 or y not in num:
            out[y] = None
        else:
            out[y] = num[y] / rev_map[y]
    return out

def series_map(standard: pd.DataFrame, stmt: str, item: str) -> Dict[int, float]:
    s = standard.loc[(standard["Statement"] == stmt) & (standard["LineItem"] == item), ["Year", "Value"]]
    return {int(y): float(v) for y, v in zip(s["Year"], s["Value"]) if pd.notna(y) and pd.notna(v)}

def write_checks_sheet(wb: Workbook, validation: ValidationResult) -> None:
    ws = wb.create_sheet("Checks")
    set_col_widths(ws, {1: 60, 2: 22, 3: 12, 4: 12, 5: 60})
    write_header_row(ws, 1, ["Check", "Year", "Diff", "Status", "Notes"])

    sc = wb._style_cache
    r = 3

    # Tolerance (for numeric diffs)
    tol = 1e-6

    if validation.checks.empty:
        ws.cell(row=r, column=1, value="No totals/tie checks could be computed from the provided dataset.")
        return

    for _, row in validation.checks.iterrows():
        diff = row.get("Diff")
        status = None if pd.isna(diff) else (abs(float(diff)) <= tol)
        ws.cell(row=r, column=1, value=str(row.get("Check", ""))).border = sc["border"]
        ws.cell(row=r, column=2, value=int(row.get("Year")) if pd.notna(row.get("Year")) else None).border = sc["border"]
        c3 = ws.cell(row=r, column=3, value=(None if pd.isna(diff) else float(diff)))
        c3.number_format = "#,##0.00"
        c3.border = sc["border"]
        ws.cell(row=r, column=4, value=("OK" if status else "FAIL")).border = sc["border"]
        ws.cell(row=r, column=5, value="Tolerance applied in code; tighten as needed.").border = sc["border"]
        r += 1

    r += 2
    # Issues table
    ws.cell(row=r, column=1, value="ROW-LEVEL ISSUES (Missing / Outliers)").fill = sc["section_fill"]
    ws.cell(row=r, column=1).font = sc["section_font"]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1

    write_header_row(ws, r, ["Statement", "LineItem", "Year", "Value", "Flags"])
    r += 2

    issues = validation.issues.copy()
    if issues.empty:
        ws.cell(row=r, column=1, value="No missing values or outliers flagged.")
        return

    for _, ir in issues.iterrows():
        flags = []
        if bool(ir.get("Missing_Flag", False)):
            flags.append("MISSING")
        if bool(ir.get("Outlier_Flag", False)):
            flags.append(f"OUTLIER(z~{ir.get('Outlier_Z'):.2f})" if pd.notna(ir.get("Outlier_Z")) else "OUTLIER")
        ws.cell(row=r, column=1, value=str(ir.get("Statement", ""))).border = sc["border"]
        ws.cell(row=r, column=2, value=str(ir.get("LineItem", ""))).border = sc["border"]
        ws.cell(row=r, column=3, value=int(ir.get("Year")) if pd.notna(ir.get("Year")) else None).border = sc["border"]
        v = ir.get("Value")
        c = ws.cell(row=r, column=4, value=(None if pd.isna(v) else float(v)))
        c.number_format = "#,##0"
        c.border = sc["border"]
        ws.cell(row=r, column=5, value="; ".join(flags)).border = sc["border"]
        r += 1

def write_forecast_stub_sheet(wb: Workbook) -> None:
    """
    Creates a Forecast sheet that matches the template layout (no numbers),
    leaving formulas to be filled by your model build later.
    """
    ws = wb.create_sheet("Forecast")
    set_col_widths(ws, {1: 40, 2: 25, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14, 11: 14, 12: 14})
    write_header_row(ws, 1, ["Line Item", "Notes", "FY-3", "FY-2", "FY-1", "FY0", "FY1", "FY2", "FY3", "FY4", "FY5", "Exit"])

    sc = wb._style_cache
    r = 3

    r = write_section(ws, r, "INCOME STATEMENT")
    for li in TEMPLATE_IS:
        ws.cell(row=r, column=1, value=li).border = sc["border"]
        ws.cell(row=r, column=2, value="").border = sc["border"]
        for col in range(3, 13):
            ws.cell(row=r, column=col, value=None).border = sc["border"]
        r += 1

    r += 1
    r = write_section(ws, r, "CASH FLOW (OPERATING MODEL)")
    cf_lines = ["EBIT", "Cash taxes on EBIT (unlevered)", "NOPAT", "D&A", "Capex", "ΔNWC", "Unlevered FCF (UFCF)"]
    for li in cf_lines:
        ws.cell(row=r, column=1, value=li).border = sc["border"]
        ws.cell(row=r, column=2, value="").border = sc["border"]
        for col in range(3, 13):
            ws.cell(row=r, column=col, value=None).border = sc["border"]
        r += 1

def write_dcf_stub_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("DCF")
    set_col_widths(ws, {1: 42, 2: 38, 3: 16, 4: 16, 5: 16, 6: 16, 7: 16, 8: 16})
    write_header_row(ws, 1, ["Item", "Formula/Notes", "FY1", "FY2", "FY3", "FY4", "FY5", "Terminal/Exit"])

    sc = wb._style_cache
    r = 3

    r = write_section(ws, r, "WACC")
    wacc_rows = [
        ("Cost of Equity", "=Risk_Free_Rate + Beta * Equity_Risk_Premium"),
        ("After-tax Cost of Debt", "=Cost_of_Debt * (1 - Tax_Rate_Marginal)"),
        ("WACC", "=(Target_Equity_%*Cost_of_Equity) + (Target_Debt_%*AfterTax_CoD)"),
    ]
    for item, note in wacc_rows:
        ws.cell(row=r, column=1, value=item).border = sc["border"]
        ws.cell(row=r, column=2, value=note).border = sc["border"]
        r += 1

    r += 1
    r = write_section(ws, r, "DISCOUNTED CASH FLOWS")
    dcf_rows = [
        ("Unlevered FCF (UFCF)", "=Forecast!UFCF"),
        ("Discount Factor", "=1/(1+WACC)^t"),
        ("PV of UFCF", "=UFCF*DiscountFactor"),
    ]
    for item, note in dcf_rows:
        ws.cell(row=r, column=1, value=item).border = sc["border"]
        ws.cell(row=r, column=2, value=note).border = sc["border"]
        for col in range(3, 8):
            ws.cell(row=r, column=col, value=None).border = sc["border"]
        ws.cell(row=r, column=8, value=None).border = sc["border"]
        r += 1

    r += 1
    r = write_section(ws, r, "TERMINAL VALUE")
    tv_rows = [
        ("TV (Gordon)", "=UFCF_(Exit+1) / (WACC - Terminal_Growth)"),
        ("TV (Exit Multiple)", "=Exit_Multiple * EBITDA_Exit"),
        ("Terminal Value (toggle)", "=IF(TV_Method_Toggle=1, TV_Gordon, TV_ExitMultiple)"),
        ("PV of Terminal Value", "=TerminalValue / (1+WACC)^ExitYear"),
        ("Enterprise Value (DCF)", "=SUM(PV_UFCF) + PV_TV"),
    ]
    for item, note in tv_rows:
        ws.cell(row=r, column=1, value=item).border = sc["border"]
        ws.cell(row=r, column=2, value=note).border = sc["border"]
        r += 1

def write_lbo_stub_sheets(wb: Workbook) -> None:
    for name in ["Debt_Schedule", "LBO_Returns", "Sensitivities", "Dashboard"]:
        ws = wb.create_sheet(name)
        write_header_row(ws, 1, ["(Template stub)", "", "", "", "", ""])
        ws.cell(row=3, column=1, value=f"{name} will be populated by your modeling layer.").font = Font(italic=True, color="666666")

# -----------------------------
# Main
# -----------------------------

def build_excel(standard: pd.DataFrame, validation: ValidationResult, out_xlsx: str) -> None:
    wb = Workbook()
    # remove default sheet
    wb.remove(wb.active)

    style_workbook(wb)

    # Create sheets in template order
    write_inputs_sheet(wb)
    write_historical_fs_sheet(wb, standard)
    write_drivers_sheet(wb, standard)
    write_forecast_stub_sheet(wb)
    write_dcf_stub_sheet(wb)
    write_lbo_stub_sheets(wb)
    write_checks_sheet(wb, validation)

    # Reorder sheets exactly (openpyxl keeps creation order, but we ensure)
    desired = ["Inputs", "Historical_FS", "Drivers", "Forecast", "DCF", "Debt_Schedule", "LBO_Returns", "Sensitivities", "Dashboard", "Checks"]
    wb._sheets = [wb[s] for s in desired if s in wb.sheetnames]

    wb.save(out_xlsx)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="in_csv", required=True, help="Path to input CSV with historical financials.")
    ap.add_argument("--outdir", required=True, help="Output directory.")
    ap.add_argument("--basename", default="model", help="Base name for outputs.")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)

    raw = read_input_csv(args.in_csv)
    long_df = to_long_form(raw)
    standard = standardize_dataset(long_df)
    validation = build_validation(standard)

    out_clean = os.path.join(args.outdir, f"{args.basename}_cleaned_standardized.csv")
    out_xlsx = os.path.join(args.outdir, f"{args.basename}_model_inputs.xlsx")

    # Save cleaned dataset
    standard_sorted = standard.sort_values(["StmtCode", "Statement", "LineRank", "LineItem", "Year"]).copy()
    standard_sorted.to_csv(out_clean, index=False)

    # Build Excel
    build_excel(standard_sorted, validation, out_xlsx)

    print(f"Saved cleaned standardized dataset: {out_clean}")
    print(f"Saved Excel template inputs:       {out_xlsx}")
    if not validation.issues.empty:
        print(f"Flagged issues rows: {len(validation.issues)} (see Checks tab).")

if __name__ == "__main__":
    main()