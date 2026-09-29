# FY 2026–27 Payroll Workbook Guide

## Main file

Use **`Automatic_Salary_Slip_PF_IT_Fixed.xlsm`** (macro-enabled). Enable macros when opening in Excel.

Legacy copies (`Automatic_Salary_Slip_PF_IT_Fixed - Copy.xlsm`, `Automatic_Salary_Slip_Final_V3.xlsm`) are kept for reference only.

## What this workbook handles

| Requirement | How it works |
|-------------|----------------|
| **PF ₹1,800 → ₹3,000 from Sep'26** | **Payroll Settings** sheet: caps in `B2` / `B3`, effective month index in `B4` (6 = Sep'26). Apr–Aug uses legacy cap; Sep–Mar uses new cap. |
| **Special allowance (not fixed)** | PF employees: `Special = HRA − PF cap` for that period. Sep+ values in **Salary sheet** columns **BT** (special) and **BU** (gross). |
| **Mid-year increment** | **Incremental sheet** table from row 9: Emp Code, Effective month, New Basic, Old Basic, Arrear months. Legacy increment layout (rows 1–5) is imported into this table. Payslip still uses existing arrear split logic in columns P–W. |
| **PDF payslips** | **Generator** sheet: pick employee & month, set output folder, run **Generate Current Slip** or **Generate All Slips**. Zero attendance for the month is skipped. |

## Sheets (quick map)

1. **Generator** – employee/month selection and PDF actions  
2. **Payroll Settings** – PF caps and Sep'26 effective month (edit here, not hard-coded in rows)  
3. **Salary sheet** – master pay, attendance (L–W), monthly payroll blocks Apr–Mar (X–BS), Sep+ components (BT–BU)  
4. **Income Tax** – annual tax and monthly TDS  
5. **Incremental sheet** – mid-year increment table  
6. **Salary Slip** – payslip view (month-aware PF and special allowance)  
7. **Master Data** – employee HR details  

## PF logic (PF Status = Yes)

- **Until Aug'26:** PF deduction = `Payroll Settings!B2` (default 1800), Special allowance = `HRA − B2`  
- **From Sep'26:** PF deduction = `Payroll Settings!B3` (default 3000), Special allowance = prior special minus `(B3 − B2)` (same as `HRA − B3` when structure is standard)  
- **PF Status = No:** PF = 0; special allowance follows column G (typically full HRA).  

Payslip **I21** (current month PF) and **J21** (cumulative PF arrears from Jul) both switch caps at the configured month.

## Rebuild after editing the “Copy” source

```bash
python3 scripts/upgrade_payroll_workbook.py
```

This reads `Automatic_Salary_Slip_PF_IT_Fixed - Copy.xlsm` and writes `Automatic_Salary_Slip_PF_IT_Fixed.xlsm`.

## Adding a new mid-year increment

1. Open **Incremental sheet**.  
2. Add a row below the table (from row 9): Emp Code, Effective month (e.g. `Oct'26`), New Basic, Old Basic, number of arrear months, Notes.  
3. Update **Salary sheet** Basic (column E) when the increment is effective.  
4. If arrear split is non-standard, adjust payslip arrear cells (P21–W21) or extend the table-driven formulas with your accountant’s rules.  

## Future PF / policy changes

Change only **Payroll Settings** (caps and effective month index). Avoid editing hundreds of employee rows for PF amounts.
