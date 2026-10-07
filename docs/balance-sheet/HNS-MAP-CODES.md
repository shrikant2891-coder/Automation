# HNS Pvt Ltd — Map Code reference (FY 2025–26)

Map Codes live in column **C** of the **TB** sheet. Notes use `SUMIF` on these codes against net balance columns **H** (CY) and **I** (PY).

## Code dictionary (current TB)

| Map Code | FS / Note area | Typical ledgers |
|----------|----------------|-----------------|
| **SC** | Share capital (Note 1) | Equity share capital per shareholder |
| **RS-OPEN** | Reserves & surplus — opening P&L | Profit & Loss A/c (opening deficit) |
| **LTB-DIR** | Long-term borrowings (Note 3) | Loans from directors |
| **TP-OTH** | Trade payables — others (Note 4) | Sundry creditors |
| **OCL-AUD** | Other current liabilities | Audit fees payable |
| **OCL-EXP** | Other current liabilities | Expenses payable |
| **LTLA-SD** | Non-current assets — LT loans & advances | Rent security deposit |
| **BANK** | Cash and cash equivalents | Bank current account |
| **CASH** | Cash and cash equivalents | Cash in hand |
| **EXP-ACC** | Other expenses (P&L) | Accounting charges |
| **EXP-AUD** | Other expenses — auditor (Note 12.1) | Statutory audit fees |
| **EXP-BANK** | Other expenses | Bank / POS charges |
| **EXP-MISC** | Other expenses | Miscellaneous |
| **EXP-OFF** | Other expenses | Office expenses |
| **EXP-ROC** | Other expenses | ROC filing fees |

## Adding a new ledger (procedure)

1. Add row on **TB** with ledger name and Tally group.
2. Set **Map Code** from table above, or define a **new code** and add a matching `SUMIF` line in **Notes** (and ratio links if needed).
3. Confirm CY/PY Dr/Cr and that net column formulas copy down.
4. Refresh FS: Balance Sheet and P&L should update without editing face cells.

## Special SUMIF patterns (not only Map Code)

Some note lines filter by **ledger name** in column A, e.g. director loans and equity split by `Lokesh*` / `Manshu*`. Use the same pattern when new directors or loan accounts are added.

## Units

- TB amounts: **rupees**
- Notes / FS presentation: often **÷ 100** → **Rupees in hundreds** on the face

Always trace one ledger from TB → Notes → Balance Sheet before sign-off.
