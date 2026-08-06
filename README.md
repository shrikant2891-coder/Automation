# RetailX Dynamic Uploader

Rebuilds the **Uploader** sheet from **Extract** + **GL Master** so new states and product categories are never missed (unlike the old hardcoded SUMIFS/VLOOKUP grid).

## Quick refresh

1. Paste the latest extract into the `Extract` sheet (keep the header row).
2. Confirm `GL Master` has GLs for every `analytics_category`.
3. Run either:

```bash
python3 build_retailx_uploader.py
```

or in Excel: `Alt+F11` → **File → Import File** → `GenerateUploader.bas` → Run macro **`GenerateUploader`**.

4. Review `Control` (debit = credit per voucher) and `Unmapped` (missing GL mappings).

## Voucher series

| Voucher | Type | Extract `report_name` |
|--------:|------|------------------------|
| 11 | Sales | `NONDIGITAL` |
| 12 | Sales Return | `RETURN_CREATED` |
| 13 | Shipping | Shipping cols on `NONDIGITAL` + `RETURN_CREATED` |
| 14 | PBO | `PBO_SALES` + `PBO_RETURN` |
| 15 | Price Drop | `PRICE_DROP` |
| 16 | Secure Packaging | `BUYER_FEE` |
| 17 | PREXO BUMPUP | `PREXO_BUMPUP` |

Change starting numbers on the `Config` sheet (documented there) / in the Python `DEFAULT_VOUCHERS` list / in the VBA constants.

## Files

- `RetailX.xlsx` — working workbook (`Uploader`, `Extract`, `GL Master`, `Config`, `Control`, `Unmapped`, `Instructions`)
- `build_retailx_uploader.py` — generator engine
- `GenerateUploader.bas` — Excel VBA equivalent
- `GG RetailX Jan'23 .xlsx` — reference workbook that used static SUMIFS
