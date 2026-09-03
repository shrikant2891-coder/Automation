# Automation

Excel automation tools for finance operations.

## MP Summary Uploader

Dynamically builds **Uploader Format** journal entries from the **Summary** pivot and **GL Backup** mapping sheet in `MP Summary.xlsx`.

### Quick start

```bash
python3 build_mp_uploader.py
```

Or with explicit paths:

```bash
python3 build_mp_uploader.py --input "MP Summary.xlsx" --output "MP Summary.xlsx"
```

### Excel macro (same approach as RetailX)

The `.xlsm` uses the **same VBA shell as `RetailX.xlsm`** with `GenerateMPUploader` swapped in.

1. Download **`MP Summary.xlsm`**
2. Open in Excel → **Enable Content**
3. `Alt+F8` → run **`GenerateMPUploader`**

**Or import manually (same as RetailX):**

1. Open `MP Summary.xlsx`
2. `Alt+F11` → **File → Import File** → select `GenerateMPUploader.bas`
3. Save as **`MP Summary.xlsm`** (macro-enabled workbook)
4. Run **`GenerateMPUploader`**

Rebuild the packaged file:

```bash
python3 package_mp_xlsm.py
```

### Voucher series

| Voucher | Type | Source |
|---------|------|--------|
| 28–29 | OI closing (Prepaid / Postpaid) | Current month `MEC-FKMP-OPEN-INVOICE-FLOW.csv` |
| 30–31 | Creditor (Postpaid / Prepaid) | Current month `FKMP-CREDITOR-REPORT (1).csv` |
| 32–33 | OI opening reversal (Prepaid / Postpaid) | **Prior month** `MEC-FKMP-OPEN-INVOICE-FLOW.csv` |
| 34 | Provision | `Provision` rows |
| 35–36 | Volume discount | `VD` / `PBO VD` rows |
| 41–44 | TCS / TDS receivable (prepaid / postpaid) | Current + prior month rows |

### Posting rules

1. Expense and GST lines are posted at **state level** using `state_code_to`.
2. TCS and TDS lines use **`tcs_state_code_to`** (fallback **`state_code_to`** when NA), include the **latest two Summary months**, and **net at state level per GL**.
3. TCS/TDS: negative net → debit receivable; positive net → credit receivable (same rule for both months).
4. GL codes are looked up from **GL Backup** using Summary column headers.
5. **IGST input**: `IN-DL` → `142067`; all other states → `142013`.
6. Debtor and provision ledgers always use **`IN-OTH`**.
7. Prior-month OI report drives **OI opening** vouchers (expense reversal).
8. **Negative** Summary values → Debit expense/GST input, Credit debtor.

After each run, check the **Control** sheet — every voucher should have `Difference = 0`.

### Other workbooks

- `RetailX.xlsm` — RetailX marketplace uploader (see `cursor/retailx-dynamic-uploader-4283` branch).
- `Automatic_Salary_Slip_Final_V3.xlsm` — salary slip automation.
