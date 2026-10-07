#!/usr/bin/env python3
"""Build Webkey comparable BS/P&L from trial balance and prior-year FS."""

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]


def r2(x) -> float:
    return float(Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP))


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

    wb = Workbook()
    ws = wb.active
    ws.title = "Balance Sheet"

    def w(row, col, val, bold=False):
        c = ws.cell(row, col, val)
        if bold:
            c.font = Font(bold=True)
        return c

    company = "WEBKEY SOLUTIONS PRIVATE LIMITED"
    cin = "U62099DL2024PTC433481"
    addr = (
        "POLE NO-58, GROUND FLOOR, VILLAGE DHANSA, NEW DELHI, "
        "South West Delhi- 110073"
    )

    w(2, 2, company, True)
    w(3, 2, f"CIN: {cin}")
    w(4, 2, "Balance Sheet as at 31 March 2026")
    w(5, 2, "(Amount in ₹ unless otherwise stated)")

    w(7, 2, "Particulars", True)
    w(7, 3, "Note", True)
    w(7, 4, "As at 31.03.2026", True)
    w(7, 5, "As at 31.03.2025", True)

    rows_bs = [
        ("I. EQUITY AND LIABILITIES", None, None, None, True),
        ("(1) Shareholders' funds", None, None, None, False),
        ("    (a) Share capital", 1, r2(share_capital), r2(py_share), False),
        ("    (b) Reserves and surplus", 2, r2(reserves), r2(py_reserves), False),
        ("", None, r2(total_equity), r2(py_equity), True),
        ("(2) Non-current liabilities", None, None, None, False),
        ("    (a) Long-term borrowings", 3, 0, 0, False),
        ("", None, 0, 0, True),
        ("(3) Current liabilities", None, None, None, False),
        ("    (a) Short-term borrowings", 4, r2(st_borrowings), 0, False),
        ("    (b) Trade payables", 5, None, None, False),
        ("        (i) total outstanding dues of micro and small enterprises", None, 0, 0, False),
        ("        (ii) total outstanding dues of creditors other than MSME", None, r2(trade_payables), 0, False),
        ("    (c) Other current liabilities", 6, r2(other_cl), 0, False),
        ("    (d) Short-term provisions", 7, r2(st_provisions), r2(py_st_provisions), False),
        ("", None, r2(total_liabilities), r2(py_total_liab), True),
        ("TOTAL", None, r2(total_le), r2(py_total_le), True),
        ("", None, None, None, False),
        ("II. ASSETS", None, None, None, True),
        ("(1) Non-current assets", None, None, None, False),
        ("    (a) Property, plant and equipment", 8, r2(ppe_net), 0, False),
        ("", None, r2(ppe_net), 0, True),
        ("(2) Current assets", None, None, None, False),
        ("    (a) Trade receivables", 9, r2(trade_receivables), 0, False),
        ("    (b) Cash and cash equivalents", 10, r2(cash_bank), r2(py_cash), False),
        ("", None, r2(trade_receivables + cash_bank), r2(py_cash), True),
        ("TOTAL", None, r2(total_assets), r2(py_total_assets), True),
        (
            "Check (Total equity and liabilities − Total assets)",
            None,
            r2(total_le - total_assets),
            r2(py_total_le - py_total_assets),
            False,
        ),
    ]

    r = 8
    for label, note, cy, py, bold in rows_bs:
        w(r, 2, label, bold)
        if note is not None:
            w(r, 3, note)
        if cy is not None:
            w(r, 4, cy)
        if py is not None:
            w(r, 5, py)
        r += 1

    ws2 = wb.create_sheet("Profit and Loss")
    for i, txt in enumerate(
        [
            company,
            f"CIN: {cin}",
            "Statement of Profit and Loss for the year ended 31 March 2026",
            "(Amount in ₹)",
        ],
        2,
    ):
        ws2.cell(i, 2, txt)
    ws2.cell(7, 2, "Particulars").font = Font(bold=True)
    ws2.cell(7, 3, "Note").font = Font(bold=True)
    ws2.cell(7, 4, "Year ended 31.03.2026").font = Font(bold=True)
    ws2.cell(7, 5, "Year ended 31.03.2025").font = Font(bold=True)

    py_rev = r2(Decimal("330001"))
    py_pat = r2(Decimal("-56680"))
    other_exp = Decimal("10000") + Decimal("4248") + Decimal("231030") + Decimal("979.66")

    ws2.cell(8, 2, "Revenue from operations")
    ws2.cell(8, 3, 11)
    ws2.cell(8, 4, r2(revenue))
    ws2.cell(8, 5, py_rev)
    ws2.cell(9, 2, "Total revenue")
    ws2.cell(9, 4, r2(revenue))
    ws2.cell(9, 5, py_rev)
    ws2.cell(11, 2, "Expenses").font = Font(bold=True)
    ws2.cell(12, 2, "Purchase / project costs")
    ws2.cell(12, 4, r2(Decimal("450000")))
    ws2.cell(13, 2, "Employee benefit expense")
    ws2.cell(13, 4, r2(Decimal("866188.8")))
    ws2.cell(14, 2, "Depreciation and amortisation")
    ws2.cell(14, 4, r2(Decimal("3132")))
    ws2.cell(15, 2, "Other expenses")
    ws2.cell(15, 4, r2(other_exp))
    ws2.cell(16, 2, "Total expenses")
    ws2.cell(16, 4, r2(expenses))
    ws2.cell(18, 2, "Profit/(loss) before tax")
    ws2.cell(18, 4, r2(pat))
    ws2.cell(18, 5, py_pat)
    ws2.cell(19, 2, "Profit/(loss) for the year")
    ws2.cell(19, 4, r2(pat))
    ws2.cell(19, 5, py_pat)

    ws3 = wb.create_sheet("TB Mapping")
    ws3.append(
        [
            "Ledger",
            "Map Code",
            "31.03.2026 Net (Dr+/Cr-)",
            "31.03.2025 Net (Dr+/Cr-)",
            "BS / P&L line",
        ]
    )
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

    ws4 = wb.create_sheet("CA Notes")
    ca_notes = [
        "1. Figures derived from TrialBaWebkeyl.xlsx (1-Apr-25 to 31-Mar-26).",
        "2. Prior year comparatives from audited Balance Sheet as at 31.03.2025 "
        "(WEBKEY SOLUTIONS PRIVATE LIMITED.xlsx); PY shows minimal operations.",
        "3. Profit & Loss A/c in Tally is a DEBIT balance (accumulated loss ₹56,679.70) — "
        "not a credit profit.",
        "4. Unsecured loans ₹45,000 classified as short-term borrowings; reclassify to "
        "long-term if terms exceed 12 months.",
        "5. MSME trade payable break-up not in TB — disclose when ageing is available.",
        "6. Income tax: assessability and provisions to be confirmed before final sign-off.",
        "7. Schedule III notes 1–18 (policies, related party, contingencies, auditor "
        "remuneration) to be completed before audit report.",
    ]
    for i, line in enumerate(ca_notes, 2):
        ws4.cell(i, 1, line)

    ws5 = wb.create_sheet("Inputs")
    for i, (k, v) in enumerate(
        [
            ("Company name", company),
            ("Registered office", addr),
            ("CIN", cin),
            ("Auditor", "G R B & CO"),
            ("FRN", "040755N"),
        ],
        2,
    ):
        ws5.cell(i, 1, k)
        ws5.cell(i, 2, v)

    out = ROOT / "WEBKEY_SOLUTIONS_FS_FY2025-26_Comparable.xlsx"
    wb.save(out)
    print(f"Wrote {out}")
    print(f"CY assets / L+E: {r2(total_assets)}")


if __name__ == "__main__":
    main()
