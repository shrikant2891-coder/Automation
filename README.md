# RetailX Dynamic Uploader

Rebuilds the **Uploader** sheet from **Extract2** + **GL Master** using explicit GST rate columns (`igst_rate`, `cgst_rate`, `sgst_utgst_rate`).

## Quick refresh

1. Paste the latest extract into the `Extract2` sheet (keep the header row).
2. Set the month on the `Config` sheet (`B2` month label, `B3` booking date).
3. Confirm `GL Master` has GLs for every `analytics_category`.
4. Run either:

```bash
python3 build_retailx_uploader.py
```

or in Excel: `Alt+F11` → run macro **`GenerateUploader`**.

5. Review `Control` (debit = credit per voucher) and `Unmapped` (missing GL mappings).

## Voucher 18 — DSS ASP provision

`SUM(Extract2[Invoice count]) × 0.385` → AR-Provision voucher (Dr 625184 / Cr 210210).

## Files

- `RetailX.xlsm` — working workbook with embedded VBA macro
- `build_retailx_uploader.py` — Python generator (uses Extract2 tax rates)
- `GenerateUploader.bas` — Excel VBA equivalent
