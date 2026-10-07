# Balance sheet workbooks — linking models & drafting guide

This folder documents how the FY **2025–26** financial statement Excel files on branch `Balance-Sheet` are structured, how numbers link between sheets, and how to draft the next year’s statements as a **Chartered Accountant** workflow.

## Workbooks in this repo

| File | Entity | Linking model | Primary use |
|------|--------|---------------|-------------|
| `HNSPVTL_FS_FY2025-26_Final.xlsx` | Private company (Schedule III) | **TB + Map Code** → Notes → FS | Companies with Tally trial balance |
| `Gupta_Agency_Financial_Statements_FY2025-26.xlsx` | Proprietorship | **Notes hub** → BS / P&L + schedules | Trading / manufacturing proprietorship |
| `Shree_Ganeshay_Traders_Financial_Statements_FY2025-26.xlsx` | Partnership | **Notes hub** + partner columns | Partnership with profit-sharing |
| `Shree_Balajee_Vertical_FS_25-26.xlsx` | Proprietorship | **Notes hub** (vertical BS) | Simpler vertical format |

## Documents

1. **[LINKING-ARCHITECTURE.md](./LINKING-ARCHITECTURE.md)** — Sheet flow, formula types (`SUMIF`, cross-sheet refs), and what to edit each year.
2. **[DRAFTING-SOP.md](./DRAFTING-SOP.md)** — CA checklist from TB/sign-off through schedules, disclosures, and review.
3. **[HNS-MAP-CODES.md](./HNS-MAP-CODES.md)** — Map Code dictionary for the company TB-driven model.

## Quick validation

From the repo root (with `openpyxl` installed):

```bash
python3 scripts/validate_balance_sheets.py
```

Checks that balance sheet totals tie (sources of funds = application of funds) for each workbook.
