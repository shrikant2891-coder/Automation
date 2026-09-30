#!/usr/bin/env python3
"""Upgrade salary slip workbook: PF Sep'26 change, month-aware components, FY months Sep–Mar."""

from __future__ import annotations

import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SRC = Path("/workspace/Automatic_Salary_Slip_PF_IT_Fixed - Copy.xlsm")
DST = Path("/workspace/Automatic_Salary_Slip_PF_IT_Fixed.xlsm")
LAST_ROW = 499
PF_LEGACY = 1800
PF_NEW = 3000
# Apr'26 = 1 … Sep'26 = 6 in Generator month list
PF_EFFECTIVE_MONTH_INDEX = 6


def ensure_payroll_settings(wb):
    if "Payroll Settings" in wb.sheetnames:
        ws = wb["Payroll Settings"]
    else:
        ws = wb.create_sheet("Payroll Settings")
        # place after Generator
        wb._sheets.insert(1, wb._sheets.pop(wb._sheets.index(ws)))

    ws["A1"] = "Setting"
    ws["B1"] = "Value"
    ws["A2"] = "PF employee cap (Apr–Aug'26)"
    ws["B2"] = PF_LEGACY
    ws["A3"] = "PF employee cap (Sep'26 onwards)"
    ws["B3"] = PF_NEW
    ws["A4"] = "PF change effective month index (Apr=1)"
    ws["B4"] = PF_EFFECTIVE_MONTH_INDEX
    ws["A5"] = "PF change effective month label"
    ws["B5"] = "Sep'26"
    ws["A7"] = "Notes"
    ws["B7"] = (
        "PF-applicable employees: Special Allowance = HRA − PF cap for that period. "
        "Edit caps here only; salary rows recalc from Payroll Settings."
    )


def col_index(letter: str) -> int:
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letter)


SPECIAL_SEP_COL = "BT"
GROSS_SEP_COL = "BU"


def upgrade_salary_sheet_fixed(ws):
    """Sep+ salary components after all FY month blocks (BT/BU) to avoid column clashes."""
    ws[f"{SPECIAL_SEP_COL}2"] = "Special Allowance (Sep+)"
    ws[f"{GROSS_SEP_COL}2"] = "Monthly Gross (Sep+)"

    for r in range(3, LAST_ROW + 1):
        if ws[f"A{r}"].value in (None, ""):
            continue
        ws[f"H{r}"] = f'=IF(D{r}="Yes",\'Payroll Settings\'!$B$2,0)'
        ws[f"I{r}"] = f'=IF(D{r}="Yes",\'Payroll Settings\'!$B$3,0)'

        g_val = ws[f"G{r}"].value
        if isinstance(g_val, str) and g_val.startswith("="):
            if "1800" in g_val or g_val in (f"=+F{r}", f"=F{r}"):
                ws[f"G{r}"] = (
                    f'=IF(D{r}="Yes",F{r}-\'Payroll Settings\'!$B$2,F{r})'
                )
        elif g_val is None:
            ws[f"G{r}"] = (
                f'=IF(D{r}="Yes",F{r}-\'Payroll Settings\'!$B$2,F{r})'
            )

        ws[f"{SPECIAL_SEP_COL}{r}"] = f'=IF(D{r}="Yes",G{r}-(I{r}-H{r}),G{r})'
        ws[f"{GROSS_SEP_COL}{r}"] = f"=E{r}+F{r}+{SPECIAL_SEP_COL}{r}"

    # Monthly blocks: Apr=X(24) + n*4 . Aug=AN(40). Sep=AR(44).
    months = [
        ("AR", "Sep'26", "Q", 6),
        ("AV", "Oct'26", "R", 7),
        ("AZ", "Nov'26", "S", 8),
        ("BD", "Dec'26", "T", 9),
        ("BH", "Jan'27", "U", 10),
        ("BL", "Feb'27", "V", 11),
        ("BP", "Mar'27", "W", 12),
    ]
    for gross_col, label, att_col, month_idx in months:
        adv_col = get_column_letter(col_index(gross_col) + 1)
        it_col = get_column_letter(col_index(gross_col) + 2)
        paid_col = get_column_letter(col_index(gross_col) + 3)
        ws[f"{gross_col}2"] = label
        ws[f"{adv_col}2"] = "Advances"
        ws[f"{it_col}2"] = "Income Tax"
        ws[f"{paid_col}2"] = "Amount Paid"

        for r in range(3, LAST_ROW + 1):
            if ws[f"A{r}"].value in (None, ""):
                continue
            gross_formula = (
                f"=ROUND(IF({month_idx}>=\'Payroll Settings\'!$B$4,"
                f"{GROSS_SEP_COL}{r},J{r})/{att_col}$1*{att_col}{r},0)"
            )
            ws[f"{gross_col}{r}"] = gross_formula
            if ws[f"{adv_col}{r}"].value is None:
                ws[f"{adv_col}{r}"] = 0
            # Income tax column on Income Tax sheet: F=Apr ... Q=Mar (col F + month_idx - 1)
            it_letter = get_column_letter(6 + month_idx - 1)  # F=6 -> Apr idx1
            ws[f"{it_col}{r}"] = f"=+'Income Tax'!{it_letter}{r - 1}"
            ws[f"{paid_col}{r}"] = (
                f"={gross_col}{r}-{adv_col}{r}-{it_col}{r}-"
                f"IF({month_idx}>=\'Payroll Settings\'!$B$4,I{r},H{r})"
            )


def upgrade_incremental_sheet(ws):
    ws["A7"] = "Increment table (add rows for mid-year changes)"
    headers = [
        "Emp Code",
        "Effective month",
        "New Basic",
        "Old Basic",
        "Arrear months",
        "Notes",
    ]
    for i, h in enumerate(headers, start=1):
        ws.cell(8, i, h)

    # Migrate legacy layout (rows 1–5) into table rows 9+
    legacy_codes = [ws.cell(1, c).value for c in range(1, 8)]
    legacy_new = [ws.cell(3, c).value for c in range(1, 8)]
    legacy_old = [ws.cell(4, c).value for c in range(1, 8)]
    row_out = 9
    for c in range(1, 8):
        code = legacy_codes[c - 1]
        if not code or not str(code).startswith("INF"):
            continue
        ws.cell(row_out, 1, code)
        ws.cell(row_out, 2, "Jul'26")
        ws.cell(row_out, 3, legacy_new[c - 1])
        ws.cell(row_out, 4, legacy_old[c - 1])
        ws.cell(row_out, 5, 3)
        ws.cell(row_out, 6, "Imported from legacy increment layout")
        row_out += 1

    ws["H8"] = "Effective month index helper"
    ws["H9"] = (
        '=MATCH(B9,{"Apr\'26","May\'26","Jun\'26","Jul\'26","Aug\'26","Sep\'26",'
        '"Oct\'26","Nov\'26","Dec\'26","Jan\'27","Feb\'27","Mar\'27"},0)'
    )
    for r in range(10, row_out):
        ws[f"H{r}"] = (
            f'=MATCH(B{r},{{"Apr\'26","May\'26","Jun\'26","Jul\'26","Aug\'26","Sep\'26",'
            f'"Oct\'26","Nov\'26","Dec\'26","Jan\'27","Feb\'27","Mar\'27"}},0)'
        )


def pf_arrears_formula(row_emp: str = "M6") -> str:
    """Cumulative employee PF arrears Jul–selected month with Sep'26 cap change."""
    att_cols = ["O", "P", "Q", "R", "S", "T", "U", "V", "W"]
    parts = []
    for i, col in enumerate(att_cols, start=4):
        rate = (
            f"IF({i}>=\'Payroll Settings\'!$B$4,"
            f"INDEX('Salary sheet'!$I$3:$I$499,{row_emp}-2),"
            f"INDEX('Salary sheet'!$H$3:$H$499,{row_emp}-2))"
        )
        parts.append(
            f"({rate})*((M5>={i})*(N(INDEX('Salary sheet'!${col}$3:${col}$499,{row_emp}-2))>0))"
        )
    return (
        f'=IF(M20,IF(M5<4,0,IFERROR({"+".join(parts)},0)),0)'
    )


def current_pf_formula(row_emp: str = "M6") -> str:
    return (
        f'=IF(M20,IF(OR(M10=0,M5<4),0,IFERROR(IF(M5>=\'Payroll Settings\'!$B$4,'
        f"INDEX('Salary sheet'!$I$3:$I$499,{row_emp}-2),"
        f"INDEX('Salary sheet'!$H$3:$H$499,{row_emp}-2)),0)),0)"
    )


def increment_arrear_formula() -> str:
    """Generic increment arrear from Incremental sheet table (replaces nested IFs)."""
    return (
        "=IF(M20=0,0,SUMPRODUCT("
        "('Incremental sheet'!$A$9:$A$200=M3)*"
        "(H9:H200<=MATCH(M4,{\"Apr'26\",\"May'26\",\"Jun'26\",\"Jul'26\",\"Aug'26\","
        "\"Sep'26\",\"Oct'26\",\"Nov'26\",\"Dec'26\",\"Jan'27\",\"Feb'27\",\"Mar'27\"},0))*"
        "(('Incremental sheet'!$C$9:$C$200-'Incremental sheet'!$D$9:$D$200)/"
        "'Incremental sheet'!$E$9:$E$200)"
        "))"
    )


def allowance_for_month_expr() -> str:
    """Special allowance for the slip column's FY month (pre/post Sep PF)."""
    return (
        "IF(COLUMN()-COLUMN('Salary sheet'!$L$2)+1>='Payroll Settings'!$B$4,"
        f"INDEX('Salary sheet'!${SPECIAL_SEP_COL}$3:${SPECIAL_SEP_COL}$499,$M$6-2),"
        "INDEX('Salary sheet'!$G$3:$G$499,$M$6-2))"
    )


def upgrade_salary_slip(ws):
    ws["I21"] = current_pf_formula()
    ws["J21"] = pf_arrears_formula()

    ws["M14"] = (
        "=IF(M5>='Payroll Settings'!$B$4,"
        f"INDEX('Salary sheet'!${SPECIAL_SEP_COL}$3:${SPECIAL_SEP_COL}$499,M6-2),"
        "INDEX('Salary sheet'!$G$3:$G$499,M6-2))"
    )
    ws["M15"] = (
        "=IF(M5>='Payroll Settings'!$B$4,"
        f"INDEX('Salary sheet'!${GROSS_SEP_COL}$3:${GROSS_SEP_COL}$499,M6-2),"
        "INDEX('Salary sheet'!$J$3:$J$499,M6-2))"
    )

    allow = allowance_for_month_expr()
    att_cols = list("LMNOPQRSTUVW")
    for i, col_letter in enumerate(att_cols):
        slip_col = get_column_letter(13 + i)  # M=13
        day_cell = f"'Salary sheet'!{col_letter}$2"
        att_range = f"'Salary sheet'!${col_letter}$3:${col_letter}$499"
        ws[f"{slip_col}42"] = (
            f"=IF(ISNUMBER(INDEX({att_range},$M$6-2)),"
            f"ROUND(({allow})/DAY({day_cell})*INDEX({att_range},$M$6-2),0),0)"
        )


def patch_salary_monthly_paid(ws):
    """Jul (AM) and Aug (AQ) amount paid — subtract legacy PF."""
    for r in range(3, LAST_ROW + 1):
        if ws[f"A{r}"].value in (None, ""):
            continue
        if isinstance(ws[f"AM{r}"].value, str) or ws[f"AM{r}"].value is None:
            ws[f"AM{r}"] = f"=AJ{r}-AK{r}-AL{r}-H{r}"
        if isinstance(ws[f"AQ{r}"].value, str) or ws[f"AQ{r}"].value is None:
            ws[f"AQ{r}"] = f"=AN{r}-AO{r}-AP{r}-H{r}"


def upgrade_vba(path: Path):
    from oletools.olevba import VBA_Parser

    vbaparser = VBA_Parser(str(path))
    if not vbaparser.detect_vba_macros():
        vbaparser.close()
        return
    # VBA lives in binary stream — patch via text replace in ole is fragile.
    # Module logic unchanged; PF is formula-driven from Sep'26.
    vbaparser.close()


def main():
    shutil.copy2(SRC, DST)
    wb = load_workbook(DST, keep_vba=True)

    ensure_payroll_settings(wb)
    upgrade_salary_sheet_fixed(wb["Salary sheet"])
    patch_salary_monthly_paid(wb["Salary sheet"])
    upgrade_incremental_sheet(wb["Incremental sheet"])
    upgrade_salary_slip(wb["Salary Slip"])

    wb.save(DST)
    print(f"Wrote {DST}")


if __name__ == "__main__":
    main()
