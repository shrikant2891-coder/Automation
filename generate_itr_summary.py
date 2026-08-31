#!/usr/bin/env python3
"""Generate ITR filing summary Excel from YES Bank statement."""

import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
STATEMENT_PATH = Path("/workspace/Account_Statement_01_Apr_2025-31_Mar_2026.xls")
OUTPUT_PATH = Path("/workspace/Arun/ITR_Summary_AY_2026-27.xlsx")

ASSESSEE = {
    "name": "ARUN SHANKAR AWASTHI",
    "pan": "",
    "email": "arun15web@gmail.com",
    "phone": "918882283917",
    "bank": "YES BANK Ltd.",
    "account_no": "014091900000507",
    "account_type": "SAVING - SMART SALARY PLATINUM",
    "branch": "BADSHAPUR, HARYANA",
    "ifsc": "YESB0000140",
    "statement_period": "01/04/2025 to 31/03/2026",
    "financial_year": "FY 2025-26",
    "assessment_year": "AY 2026-27",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
LABEL_FONT = Font(bold=True)
MONEY_FMT = "#,##0.00"
DATE_FMT = "DD-MMM-YYYY"
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def parse_amount(value: str) -> float:
    if not value:
        return 0.0
    return float(value.replace(",", ""))


def parse_date(value: str) -> datetime | None:
    try:
        return datetime.strptime(value, "%d %b %Y")
    except ValueError:
        return None


def get_row_values(row) -> dict[int, str]:
    cells = row.findall("ss:Cell", NS)
    result: dict[int, str] = {}
    col = 0
    for cell in cells:
        idx = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
        if idx:
            col = int(idx) - 1
        data = cell.find("ss:Data", NS)
        val = (data.text or "").strip() if data is not None else ""
        result[col] = val
        col += 1
    return result


def load_transactions() -> list[dict]:
    root = ET.fromstring(STATEMENT_PATH.read_text(encoding="utf-8"))
    rows = root.findall(".//ss:Worksheet/ss:Table/ss:Row", NS)
    header_idx = next(
        i for i, row in enumerate(rows) if get_row_values(row).get(0) == "Transaction Date"
    )

    transactions = []
    for row in rows[header_idx + 1 :]:
        vals = get_row_values(row)
        if not re.match(r"\d{2} \w{3} \d{4}", vals.get(0, "")):
            continue
        txn_date = parse_date(vals.get(0, ""))
        value_date = parse_date(vals.get(1, ""))
        transactions.append(
            {
                "txn_date": txn_date,
                "txn_date_str": vals.get(0, ""),
                "value_date": value_date,
                "value_date_str": vals.get(1, ""),
                "ref_no": vals.get(2, ""),
                "description": vals.get(3, ""),
                "withdrawal": parse_amount(vals.get(4, "")),
                "deposit": parse_amount(vals.get(5, "")),
                "balance": parse_amount(vals.get(6, "")),
            }
        )
    return transactions


def is_infutive_salary(txn: dict) -> bool:
    desc = txn["description"].lower()
    return "infutive" in desc and txn["deposit"] > 0 and "salary" in desc


def is_infutive_related(txn: dict) -> bool:
    return "infutive" in txn["description"].lower()


def is_interest_income(txn: dict) -> bool:
    return "interest" in txn["description"].lower() and "capitalised" in txn["description"].lower()


def is_cash_deposit(txn: dict) -> bool:
    desc = txn["description"].upper()
    return ("CASH DEPOSIT" in desc or "BNA CASH" in desc or "CHQ DEPOSIT" in desc) and txn["deposit"] > 0


def classify_income(txn: dict) -> str:
    desc = txn["description"].upper()
    if is_infutive_salary(txn):
        return "Salary - Infutive Technology"
    if is_interest_income(txn):
        return "Interest Income (Savings Bank)"
    if "PAYTM MONEY" in desc:
        return "Investment Redemption (Paytm Money)"
    if "CASH DEPOSIT" in desc or "BNA CASH" in desc:
        return "Cash Deposit"
    if "CHQ DEPOSIT" in desc:
        return "Cheque Deposit"
    if txn["deposit"] > 0 and ("NEFT CR" in desc or "/FROM:" in desc):
        return "Other Credits (Transfers/Receipts)"
    return "Other"


def style_header_row(ws, row_num: int, col_count: int) -> None:
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def write_table(
    ws,
    start_row: int,
    headers: list[str],
    rows: list[list],
    money_cols: set[int] | None = None,
    date_cols: set[int] | None = None,
) -> int:
    money_cols = money_cols or set()
    date_cols = date_cols or set()

    for col, header in enumerate(headers, 1):
        ws.cell(row=start_row, column=col, value=header)
    style_header_row(ws, start_row, len(headers))

    for r_idx, row in enumerate(rows, start_row + 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c_idx in money_cols and isinstance(value, (int, float)):
                cell.number_format = MONEY_FMT
            if c_idx in date_cols and isinstance(value, datetime):
                cell.number_format = DATE_FMT

    return start_row + len(rows)


def autosize_columns(ws, min_width: int = 10, max_width: int = 55) -> None:
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        length = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[letter].width = min(max(length + 2, min_width), max_width)


def create_itr_summary_sheet(wb: Workbook, transactions: list[dict]) -> None:
    ws = wb.active
    ws.title = "ITR Summary"

    salary_txns = [t for t in transactions if is_infutive_salary(t)]
    interest_txns = [t for t in transactions if is_interest_income(t)]
    cash_txns = [t for t in transactions if is_cash_deposit(t)]

    total_salary = sum(t["deposit"] for t in salary_txns)
    total_interest = sum(t["deposit"] for t in interest_txns)
    total_cash_deposits = sum(t["deposit"] for t in cash_txns)
    total_deposits = sum(t["deposit"] for t in transactions)
    total_withdrawals = sum(t["withdrawal"] for t in transactions)

    paytm_redemptions = [
        t for t in transactions if "paytm money" in t["description"].lower() and t["deposit"] > 0
    ]
    total_paytm = sum(t["deposit"] for t in paytm_redemptions)

    row = 1
    ws.cell(row=row, column=1, value="ITR Filing Summary - Bank Statement Analysis").font = TITLE_FONT
    row += 1
    ws.cell(row=row, column=1, value=f"{ASSESSEE['financial_year']} | {ASSESSEE['assessment_year']}")
    row += 2

    details = [
        ("Assessee Name", ASSESSEE["name"]),
        ("Email", ASSESSEE["email"]),
        ("Phone", ASSESSEE["phone"]),
        ("Bank", ASSESSEE["bank"]),
        ("Account Number", ASSESSEE["account_no"]),
        ("Account Type", ASSESSEE["account_type"]),
        ("Branch", ASSESSEE["branch"]),
        ("IFSC", ASSESSEE["ifsc"]),
        ("Statement Period", ASSESSEE["statement_period"]),
        ("Total Transactions", len(transactions)),
    ]
    for label, value in details:
        ws.cell(row=row, column=1, value=label).font = LABEL_FONT
        ws.cell(row=row, column=2, value=value)
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Income Summary (for ITR)").font = TITLE_FONT
    row += 1

    income_rows = [
        ["Income Head", "Amount (INR)", "No. of Entries", "ITR Schedule / Notes"],
        [
            "Salary from Infutive Technology Pvt Ltd",
            total_salary,
            len(salary_txns),
            "Schedule S - See 'Infutive Salary' sheet",
        ],
        [
            "Interest from Savings Bank Account",
            total_interest,
            len(interest_txns),
            "Schedule OS / Section 80TTA (exempt up to Rs.10,000)",
        ],
        [
            "Investment Redemption (Paytm Money)",
            total_paytm,
            len(paytm_redemptions),
            "Schedule CG / OS - Capital gains or other sources as applicable",
        ],
        [
            "Cash Deposits (Self/Third Party)",
            total_cash_deposits,
            len(cash_txns),
            "Disclose in ITR cash transaction schedule if applicable",
        ],
        [
            "Gross Total Credits in Account",
            total_deposits,
            sum(1 for t in transactions if t["deposit"] > 0),
            "Not all credits are taxable income",
        ],
    ]
    end_row = write_table(ws, row, income_rows[0], income_rows[1:], money_cols={2}, date_cols=set())
    row = end_row + 2

    ws.cell(row=row, column=1, value="Key Figures").font = TITLE_FONT
    row += 1
    key_rows = [
        ["Particulars", "Amount (INR)"],
        ["Total Deposits (Credits)", total_deposits],
        ["Total Withdrawals (Debits)", total_withdrawals],
        ["Net Flow (Deposits - Withdrawals)", total_deposits - total_withdrawals],
        ["Closing Balance (as per last txn)", transactions[0]["balance"] if transactions else 0],
        ["Taxable Salary Identified (Infutive)", total_salary],
        ["Savings Bank Interest", total_interest],
    ]
    write_table(ws, row, key_rows[0], key_rows[1:], money_cols={2})

    row += len(key_rows) + 2
    ws.cell(row=row, column=1, value="Important Notes for ITR Filing").font = TITLE_FONT
    notes = [
        "1. Salary entries are identified from Infutive Technology Private Limited UPI credits marked 'Salary'.",
        "2. Only 3 salary credits found in Aug 2025 (Rs.5,000 each). Verify with Form 16 / salary slips.",
        "3. Savings bank interest of Rs.758 is eligible for deduction u/s 80TTA (up to Rs.10,000).",
        "4. Large credits from individuals/businesses are transfers, not salary. Review before declaring as income.",
        "5. IMPS debits to Infutive Technology (Feb 2026) appear to be loan repayments - not salary.",
        "6. Cash deposits should be reconciled with source of funds for ITR disclosure requirements.",
        "7. Paytm Money NEFT credits may relate to mutual fund/stock redemptions - check capital gains.",
    ]
    for note in notes:
        row += 1
        ws.cell(row=row, column=1, value=note)

    autosize_columns(ws)


def create_infutive_salary_sheet(wb: Workbook, transactions: list[dict]) -> None:
    ws = wb.create_sheet("Infutive Salary")
    salary_txns = sorted(
        [t for t in transactions if is_infutive_salary(t)],
        key=lambda t: t["txn_date"] or datetime.min,
    )

    row = 1
    ws.cell(row=row, column=1, value="Salary from Infutive Technology Private Limited").font = TITLE_FONT
    row += 1
    ws.cell(
        row=row,
        column=1,
        value="Credits from employer to be reported under 'Income from Salary' (Schedule S)",
    )
    row += 2

    headers = [
        "S.No.",
        "Transaction Date",
        "Value Date",
        "Reference No.",
        "Description",
        "Salary Amount (INR)",
        "Running Balance",
        "Remarks",
    ]
    data = []
    for idx, txn in enumerate(salary_txns, 1):
        data.append(
            [
                idx,
                txn["txn_date"],
                txn["value_date"],
                txn["ref_no"],
                txn["description"],
                txn["deposit"],
                txn["balance"],
                "UPI credit marked as Salary",
            ]
        )

    total = sum(t["deposit"] for t in salary_txns)
    data.append(["", "", "", "", "TOTAL SALARY (Infutive Technology)", total, "", ""])

    end_row = write_table(
        ws,
        row,
        headers,
        data,
        money_cols={6, 7},
        date_cols={2, 3},
    )

    total_row = end_row
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=total_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = SUBHEADER_FILL

    row = total_row + 3
    ws.cell(row=row, column=1, value="Other Infutive Technology Entries (Not Salary)").font = TITLE_FONT
    row += 1

    other_headers = [
        "Transaction Date",
        "Type",
        "Amount (INR)",
        "Description",
        "Suggested Treatment",
    ]
    other_data = []
    for txn in sorted(
        [t for t in transactions if is_infutive_related(t) and not is_infutive_salary(t)],
        key=lambda t: t["txn_date"] or datetime.min,
    ):
        if txn["withdrawal"] > 0:
            amount = txn["withdrawal"]
            txn_type = "Debit (Payment to Infutive)"
            treatment = "Loan repayment / transfer - not salary income"
        else:
            amount = txn["deposit"]
            txn_type = "Credit"
            treatment = "Review manually"
        other_data.append(
            [
                txn["txn_date"],
                txn_type,
                amount,
                txn["description"],
                treatment,
            ]
        )

    write_table(ws, row, other_headers, other_data, money_cols={3}, date_cols={1})
    autosize_columns(ws)


def create_interest_sheet(wb: Workbook, transactions: list[dict]) -> None:
    ws = wb.create_sheet("Interest Income")
    interest_txns = sorted(
        [t for t in transactions if is_interest_income(t)],
        key=lambda t: t["txn_date"] or datetime.min,
    )

    headers = ["S.No.", "Date", "Description", "Interest Amount (INR)", "Quarter"]
    rows = []
    for idx, txn in enumerate(interest_txns, 1):
        quarter = ""
        if txn["txn_date"]:
            month = txn["txn_date"].month
            if month in (4, 5, 6):
                quarter = "Q1 (Apr-Jun)"
            elif month in (7, 8, 9):
                quarter = "Q2 (Jul-Sep)"
            elif month in (10, 11, 12):
                quarter = "Q3 (Oct-Dec)"
            else:
                quarter = "Q4 (Jan-Mar)"
        rows.append([idx, txn["txn_date"], txn["description"], txn["deposit"], quarter])

    total = sum(t["deposit"] for t in interest_txns)
    rows.append(["", "", "TOTAL", total, "Deduction u/s 80TTA available"])

    write_table(ws, 1, headers, rows, money_cols={4}, date_cols={2})
    autosize_columns(ws)


def create_cash_deposits_sheet(wb: Workbook, transactions: list[dict]) -> None:
    ws = wb.create_sheet("Cash Deposits")
    cash_txns = sorted(
        [t for t in transactions if is_cash_deposit(t)],
        key=lambda t: t["txn_date"] or datetime.min,
    )

    headers = [
        "S.No.",
        "Transaction Date",
        "Reference No.",
        "Description",
        "Deposit Amount (INR)",
        "Source to Declare in ITR",
    ]
    rows = []
    for idx, txn in enumerate(cash_txns, 1):
        rows.append(
            [
                idx,
                txn["txn_date"],
                txn["ref_no"],
                txn["description"],
                txn["deposit"],
                "Self / Business / Other - verify source",
            ]
        )
    total = sum(t["deposit"] for t in cash_txns)
    rows.append(["", "", "", "TOTAL CASH DEPOSITS", total, ""])

    write_table(ws, 1, headers, rows, money_cols={5}, date_cols={2})
    autosize_columns(ws)


def create_monthly_summary_sheet(wb: Workbook, transactions: list[dict]) -> None:
    ws = wb.create_sheet("Monthly Summary")
    monthly: dict[str, dict] = defaultdict(lambda: {"deposits": 0.0, "withdrawals": 0.0, "count": 0})

    for txn in transactions:
        if not txn["txn_date"]:
            continue
        key = txn["txn_date"].strftime("%b-%Y")
        monthly[key]["deposits"] += txn["deposit"]
        monthly[key]["withdrawals"] += txn["withdrawal"]
        monthly[key]["count"] += 1

    def month_sort_key(item: str) -> datetime:
        return datetime.strptime(item, "%b-%Y")

    headers = ["Month", "No. of Txns", "Total Deposits (INR)", "Total Withdrawals (INR)", "Net (INR)"]
    rows = []
    for month in sorted(monthly.keys(), key=month_sort_key):
        data = monthly[month]
        rows.append(
            [
                month,
                data["count"],
                data["deposits"],
                data["withdrawals"],
                data["deposits"] - data["withdrawals"],
            ]
        )

    write_table(ws, 1, headers, rows, money_cols={3, 4, 5})
    autosize_columns(ws)


def create_all_transactions_sheet(wb: Workbook, transactions: list[dict]) -> None:
    ws = wb.create_sheet("All Transactions")
    headers = [
        "S.No.",
        "Transaction Date",
        "Value Date",
        "Reference No.",
        "Description",
        "Withdrawal (INR)",
        "Deposit (INR)",
        "Balance (INR)",
        "Category",
    ]

    sorted_txns = sorted(transactions, key=lambda t: t["txn_date"] or datetime.min, reverse=True)
    rows = []
    for idx, txn in enumerate(sorted_txns, 1):
        category = classify_income(txn) if txn["deposit"] > 0 else "Expense/Transfer"
        rows.append(
            [
                idx,
                txn["txn_date"],
                txn["value_date"],
                txn["ref_no"],
                txn["description"],
                txn["withdrawal"] or None,
                txn["deposit"] or None,
                txn["balance"],
                category,
            ]
        )

    write_table(ws, 1, headers, rows, money_cols={6, 7, 8}, date_cols={2, 3})
    autosize_columns(ws)


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    transactions = load_transactions()

    wb = Workbook()
    create_itr_summary_sheet(wb, transactions)
    create_infutive_salary_sheet(wb, transactions)
    create_interest_sheet(wb, transactions)
    create_cash_deposits_sheet(wb, transactions)
    create_monthly_summary_sheet(wb, transactions)
    create_all_transactions_sheet(wb, transactions)

    wb.save(OUTPUT_PATH)
    print(f"Created: {OUTPUT_PATH}")
    print(f"Transactions processed: {len(transactions)}")
    print(f"Infutive salary entries: {len([t for t in transactions if is_infutive_salary(t)])}")
    print(f"Total salary amount: {sum(t['deposit'] for t in transactions if is_infutive_salary(t)):,.2f}")


if __name__ == "__main__":
    main()
