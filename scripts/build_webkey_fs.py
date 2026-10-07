#!/usr/bin/env python3
"""Build Webkey comparable BS/P&L from trial balance and prior-year FS."""

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]

INR_FMT = '#,##0.00;(#,##0.00);"-"'
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="808080")


def r2(x) -> float:
    return float(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


def style_header_row(ws, row: int, col_start: int, col_end: int) -> None:
    for col in range(col_start, col_end + 1):
        c = ws.cell(row, col)
        c.font = Font(bold=True, size=11)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(THIN, THIN, THIN, THIN)


def style_amount_cols(ws, row_start: int, row_end: int, cols: tuple[int, ...]) -> None:
    for row in range(row_start, row_end + 1):
        for col in cols:
            c = ws.cell(row, col)
            if isinstance(c.value, (int, float)):
                c.number_format = INR_FMT
                c.alignment = Alignment(horizontal="right", vertical="center")


def apply_table_border(ws, row_start: int, row_end: int, col_start: int, col_end: int) -> None:
    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            ws.cell(row, col).border = Border(THIN, THIN, THIN, THIN)


def setup_sheet_layout(ws, particulars_width=52, note_width=8, amt_width=18) -> None:
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = particulars_width
    ws.column_dimensions["C"].width = note_width
    ws.column_dimensions["D"].width = amt_width
    ws.column_dimensions["E"].width = amt_width
    ws.sheet_view.showGridLines = True


def main() -> None:
    share_capital = Decimal("50000")
    pl_bf_debit = Decimal("56679.7")
    revenue = Decimal("1526000") + Decimal("1")
    expenses = (
        Decimal("450000")
        + Decimal("10000")
        + Decimal("4248")
        + Decimal("3132")
        + Decimal("231030")
        + Decimal("866188.8")
        + Decimal("979.66")
    )
    pat = revenue - expenses
    reserves = -pl_bf_debit + pat
    total_equity = share_capital + reserves

    st_borrowings = Decimal("45000")
    trade_payables = Decimal("621146")
    other_cl = Decimal("31149.91") + Decimal("20000")
    st_provisions = Decimal("0")
    total_liabilities = st_borrowings + trade_payables + other_cl + st_provisions
    total_le = total_equity + total_liabilities

    ppe_net = Decimal("75415.25") - Decimal("3132")
    trade_receivables = Decimal("77000")
    cash_bank = Decimal("521755.5")
    total_assets = ppe_net + trade_receivables + cash_bank

    py_share = Decimal("50000")
    py_reserves = Decimal("-56679.6")
    py_equity = py_share + py_reserves
    py_st_provisions = Decimal("10000")
    py_total_liab = py_st_provisions
    py_total_le = py_equity + py_total_liab
    py_cash = Decimal("3320.4")
    py_total_assets = py_cash

    if abs(total_assets - total_le) > Decimal("0.02"):
        raise SystemExit(f"BS does not tie: assets {total_assets} vs L+E {total_le}")

    company = "WEBKEY SOLUTIONS PRIVATE LIMITED"
    cin = "U62099DL2024PTC433481"
    addr = (
        "POLE NO-58, GROUND FLOOR, VILLAGE DHANSA, NEW DELHI, "
        "South West Delhi- 110073"
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Balance Sheet"
    setup_sheet_layout(ws)

    title_font = Font(bold=True, size=14)
    sub_font = Font(size=11)

    ws.merge_cells("B2:E2")
    c = ws["B2"]
    c.value = company
    c.font = title_font
    c.alignment = Alignment(horizontal="center")

    ws.merge_cells("B3:E3")
    ws["B3"].value = f"CIN: {cin}"
    ws["B3"].font = sub_font
    ws["B3"].alignment = Alignment(horizontal="center")

    ws.merge_cells("B4:E4")
    ws["B4"].value = "Balance Sheet as at 31 March 2026"
    ws["B4"].font = Font(bold=True, size=12)
    ws["B4"].alignment = Alignment(horizontal="center")

    ws.merge_cells("B5:E5")
    ws["B5"].value = "(Amount in ₹ unless otherwise stated)"
    ws["B5"].font = Font(italic=True, size=10)
    ws["B5"].alignment = Alignment(horizontal="center")

    header_row = 7
    ws.cell(header_row, 2, "Particulars")
    ws.cell(header_row, 3, "Note No.")
    ws.cell(header_row, 4, "As at\n31 March 2026")
    ws.cell(header_row, 5, "As at\n31 March 2025")
    style_header_row(ws, header_row, 2, 5)
    ws.row_dimensions[header_row].height = 32

    rows_bs = [
        ("I. EQUITY AND LIABILITIES", None, None, None, True, False),
        ("(1) Shareholders' funds", None, None, None, False, False),
        ("    (a) Share capital", 1, r2(share_capital), r2(py_share), False, False),
        ("    (b) Reserves and surplus", 2, r2(reserves), r2(py_reserves), False, False),
        ("", None, r2(total_equity), r2(py_equity), True, True),
        ("(2) Non-current liabilities", None, None, None, False, False),
        ("    (a) Long-term borrowings", 3, 0, 0, False, False),
        ("", None, 0, 0, True, True),
        ("(3) Current liabilities", None, None, None, False, False),
        ("    (a) Short-term borrowings", 4, r2(st_borrowings), 0, False, False),
        ("    (b) Trade payables", 5, None, None, False, False),
        (
            "        (i) Total outstanding dues of micro and small enterprises",
            None,
            0,
            0,
            False,
            False,
        ),
        (
            "        (ii) Total outstanding dues of creditors other than MSME",
            None,
            r2(trade_payables),
            0,
            False,
            False,
        ),
        ("    (c) Other current liabilities", 6, r2(other_cl), 0, False, False),
        ("    (d) Short-term provisions", 7, r2(st_provisions), r2(py_st_provisions), False, False),
        ("", None, r2(total_liabilities), r2(py_total_liab), True, True),
        ("TOTAL", None, r2(total_le), r2(py_total_le), True, True),
        ("", None, None, None, False, False),
        ("II. ASSETS", None, None, None, True, False),
        ("(1) Non-current assets", None, None, None, False, False),
        ("    (a) Property, plant and equipment", 8, r2(ppe_net), 0, False, False),
        ("", None, r2(ppe_net), 0, True, True),
        ("(2) Current assets", None, None, None, False, False),
        ("    (a) Trade receivables", 9, r2(trade_receivables), 0, False, False),
        ("    (b) Cash and cash equivalents", 10, r2(cash_bank), r2(py_cash), False, False),
        ("", None, r2(trade_receivables + cash_bank), r2(py_cash), True, True),
        ("TOTAL", None, r2(total_assets), r2(py_total_assets), True, True),
        (
            "Check (Total equity and liabilities − Total assets)",
            None,
            r2(total_le - total_assets),
            r2(py_total_le - py_total_assets),
            False,
            False,
        ),
    ]

    r = 8
    total_rows = []
    for label, note, cy, py, bold, subtotal in rows_bs:
        cell_b = ws.cell(r, 2, label)
        cell_b.alignment = Alignment(vertical="center", wrap_text=True)
        if bold:
            cell_b.font = Font(bold=True)
        if note is not None:
            ws.cell(r, 3, note).alignment = Alignment(horizontal="center")
        if cy is not None:
            ws.cell(r, 4, cy)
        if py is not None:
            ws.cell(r, 5, py)
        if subtotal or label == "TOTAL":
            total_rows.append(r)
            for col in range(2, 6):
                ws.cell(r, col).fill = TOTAL_FILL
        r += 1

    last_row = r - 1
    style_amount_cols(ws, 8, last_row, (4, 5))
    apply_table_border(ws, header_row, last_row, 2, 5)
    ws.freeze_panes = "B8"

    # --- P&L ---
    ws2 = wb.create_sheet("Profit and Loss")
    setup_sheet_layout(ws2)
    for merge_row, txt, font in [
        (2, company, title_font),
        (3, f"CIN: {cin}", sub_font),
        (4, "Statement of Profit and Loss for the year ended 31 March 2026", Font(bold=True, size=12)),
        (5, "(Amount in ₹)", Font(italic=True, size=10)),
    ]:
        ws2.merge_cells(f"B{merge_row}:E{merge_row}")
        ws2.cell(merge_row, 2, txt).font = font
        ws2.cell(merge_row, 2).alignment = Alignment(horizontal="center")

    pl_header = 7
    ws2.cell(pl_header, 2, "Particulars")
    ws2.cell(pl_header, 3, "Note No.")
    ws2.cell(pl_header, 4, "Year ended\n31 March 2026")
    ws2.cell(pl_header, 5, "Year ended\n31 March 2025")
    style_header_row(ws2, pl_header, 2, 5)
    ws2.row_dimensions[pl_header].height = 32

    py_rev = r2(Decimal("330001"))
    py_pat = r2(Decimal("-56680"))
    other_exp = Decimal("10000") + Decimal("4248") + Decimal("231030") + Decimal("979.66")

    pl_data = [
        ("Revenue from operations", 11, r2(revenue), py_rev, False),
        ("Total revenue", None, r2(revenue), py_rev, True),
        ("", None, None, None, False),
        ("Expenses", None, None, None, True),
        ("Purchase / project costs", None, r2(Decimal("450000")), None, False),
        ("Employee benefit expense", None, r2(Decimal("866188.8")), None, False),
        ("Depreciation and amortisation", None, r2(Decimal("3132")), None, False),
        ("Other expenses", None, r2(other_exp), None, False),
        ("Total expenses", None, r2(expenses), None, True),
        ("", None, None, None, False),
        ("Profit/(loss) before tax", None, r2(pat), py_pat, True),
        ("Profit/(loss) for the year", None, r2(pat), py_pat, True),
    ]
    pr = 8
    for label, note, cy, py, bold in pl_data:
        ws2.cell(pr, 2, label)
        if bold:
            ws2.cell(pr, 2).font = Font(bold=True)
        if note:
            ws2.cell(pr, 3, note)
        if cy is not None:
            ws2.cell(pr, 4, cy)
        if py is not None:
            ws2.cell(pr, 5, py)
        if bold:
            for col in range(2, 6):
                ws2.cell(pr, col).fill = TOTAL_FILL
        pr += 1
    style_amount_cols(ws2, 8, pr - 1, (4, 5))
    apply_table_border(ws2, pl_header, pr - 1, 2, 5)
    ws2.freeze_panes = "B8"

    # --- TB Mapping ---
    ws3 = wb.create_sheet("TB Mapping")
    ws3.column_dimensions["A"].width = 36
    ws3.column_dimensions["B"].width = 12
    for col in "CDEF":
        ws3.column_dimensions[col].width = 20
    ws3.column_dimensions["E"].width = 28
    headers = [
        "Ledger",
        "Map Code",
        "31.03.2026 Net (Dr+/Cr-)",
        "31.03.2025 Net (Dr+/Cr-)",
        "BS / P&L line",
    ]
    ws3.append(headers)
    style_header_row(ws3, 1, 1, 5)
    for row in [
        ("Share capital (Lokesh + Ritu)", "SC", -50000, -50000, "BS Share capital"),
        ("Profit & Loss A/c (debit b/f)", "RS-OPEN", 56679.7, 56679.7, "Reserves (opening loss)"),
        ("Unsecured loans", "STB", -45000, 0, "ST borrowings"),
        ("Sundry creditors", "TP-OTH", -621146, 0, "Trade payables"),
        ("Duties & taxes + audit payable", "OCL", -51149.91, -10000, "Other CL / provisions"),
        ("PPE net", "PPE", r2(ppe_net), 0, "PPE"),
        ("Trade receivables", "TR", 77000, 0, "Trade receivables"),
        ("Bank", "CASH", 521755.5, 3320.3, "Cash"),
        ("IT Services", "REV", -1526000, -330001, "Revenue"),
        ("Design & printing", "COGS", 450000, 0, "Purchase"),
    ]:
        ws3.append(list(row))
    style_amount_cols(ws3, 2, ws3.max_row, (3, 4))
    apply_table_border(ws3, 1, ws3.max_row, 1, 5)
    ws3.freeze_panes = "A2"

    # --- CA Notes ---
    ws4 = wb.create_sheet("CA Notes")
    ws4.column_dimensions["A"].width = 110
    ws4["A1"] = "Chartered Accountant — review notes (draft)"
    ws4["A1"].font = Font(bold=True, size=12)
    ca_notes = [
        "1. Figures derived from TrialBaWebkeyl.xlsx (1-Apr-25 to 31-Mar-26).",
        "2. Prior year comparatives from audited Balance Sheet as at 31.03.2025.",
        "3. Profit & Loss A/c in Tally is a DEBIT balance (accumulated loss ₹56,679.70).",
        "4. Unsecured loans ₹45,000 classified as short-term borrowings pending loan agreement review.",
        "5. MSME trade payable break-up to be added when ageing is available.",
        "6. Income tax and full Schedule III notes 1–18 to be completed before audit report and UDIN.",
    ]
    for i, line in enumerate(ca_notes, 3):
        ws4.cell(i, 1, line)
        ws4.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="top")

    # --- Inputs ---
    ws5 = wb.create_sheet("Inputs")
    ws5.column_dimensions["A"].width = 22
    ws5.column_dimensions["B"].width = 70
    ws5["A1"] = "Field"
    ws5["B1"] = "Value"
    style_header_row(ws5, 1, 1, 2)
    for i, (k, v) in enumerate(
        [
            ("Company name", company),
            ("Registered office", addr),
            ("CIN", cin),
            ("Financial year", "2025-26 (year ended 31 March 2026)"),
            ("Auditor", "G R B & CO"),
            ("FRN", "040755N"),
            ("Source TB", "TrialBaWebkeyl.xlsx"),
            ("Source PY FS", "WEBKEY SOLUTIONS PRIVATE LIMITED.xlsx"),
        ],
        2,
    ):
        ws5.cell(i, 1, k)
        ws5.cell(i, 2, v)
        ws5.cell(i, 2).alignment = Alignment(wrap_text=True)
    apply_table_border(ws5, 1, 9, 1, 2)

    out = ROOT / "WEBKEY_SOLUTIONS_FS_FY2025-26_Comparable.xlsx"
    wb.save(out)

    artifacts = Path("/opt/cursor/artifacts")
    if artifacts.is_dir():
        import shutil

        shutil.copy2(out, artifacts / out.name)

    print(f"Wrote {out}")
    print(f"CY assets / L+E: {r2(total_assets)}")


if __name__ == "__main__":
    main()
