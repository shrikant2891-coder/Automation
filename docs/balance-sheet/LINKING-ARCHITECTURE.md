# Balance sheet linking architecture

Two families of templates are in use. Choose the family by **entity type** and whether you have a **mapped trial balance (TB)**.

```
                    ┌─────────────────────────────────────────┐
                    │           FAMILY A: TB-driven (Co.)      │
                    │  Tally TB → Map Code → Notes (SUMIF)   │
                    │       → Balance Sheet / P&L / Cash Flow  │
                    │       ← Inputs (entity + auditor meta)   │
                    └─────────────────────────────────────────┘

                    ┌─────────────────────────────────────────┐
                    │     FAMILY B: Notes-hub (Prop. / Firm)   │
                    │  Schedules + manual/TB in Notes rows     │
                    │       → Balance Sheet / P&L (cell refs)  │
                    │  P&L net profit → Notes (capital rollfwd)│
                    └─────────────────────────────────────────┘
```

---

## Family A — `HNSPVTL_FS_FY2025-26_Final.xlsx` (company)

### Sheet roles

| Sheet | Role |
|-------|------|
| **TB** | Source of truth: each ledger has **Map Code**, Dr/Cr for CY and PY. Net balance columns (H/I) feed Notes. |
| **Inputs** | Yellow cells: company name, CIN, registered office, share count, auditor, DINs, place/date. Headers on BS/P&L/CF pull from here. |
| **Notes** | Disclosure tables; amounts from `SUMIF(TB!$C:$C,"<code>",TB!$H:$H)` (and PY column I), often `/100` because FS are in **hundreds**. |
| **Balance Sheet** | Schedule III line items = **direct links to Notes totals** (e.g. share capital = `Notes!C10`). |
| **Profit and Loss** | Same pattern from Notes (revenue, expenses). |
| **Cash Flow** | Indirect method: pulls movement from BS and P&L lines. |
| **Ratios** | Schedule III analytical ratios from BS + P&L numerators/denominators. |
| **Queries & Assumptions** | Working queries (not part of published FS). |

### Linking types (formulas)

| Type | Example | Purpose |
|------|---------|---------|
| **SUMIF on Map Code** | `=SUMIF(TB!$C:$C,"SC",TB!$H:$H)` | Roll all ledgers tagged `SC` into share capital note. |
| **Ledger name SUMIF** | `SUMIF(TB!$A:$A,"Lokesh*Equity*",...)` | Split director loan vs equity when codes alone are not enough. |
| **Notes → FS** | `=Notes!C37` on Balance Sheet | Single note total drives face of FS. |
| **Inputs → header** | `=Inputs!$B$4` | One place to update entity name. |
| **SUM subtotals** | `=SUM(C9:C9)` in Notes | Note section totals before FS link. |
| **Check row** | `Total liabilities − Total assets` | Must be **0**. |

### TB column logic (conceptual)

- **Map Code** (column C): stable tag — do not change codes on FS sheets; only on TB.
- **CY net**: Dr − Cr (or dedicated net column) for 31.03.2026.
- **PY net**: prior year column for comparatives.

When you add a new ledger in Tally export, you **only** assign a Map Code on TB; Notes and BS update if the code already exists in `SUMIF` ranges.

### Unit convention

- TB is in **rupees**; many Notes formulas divide by **100** so the face of FS shows **Rupees in hundreds** (Schedule III practice).

---

## Family B — Proprietorship / partnership (Gupta, Ganeshay, Balajee)

### Sheet roles

| Sheet | Role |
|-------|------|
| **Notes** | **Hub**: line-item detail, section subtotals, and links from P&L for capital/profit. |
| **Balance Sheet** | Each line = `=Notes!<cell>` for CY and PY columns; **TOTAL** rows sum BS lines only. |
| **Profit & Loss** | Expense/income lines often `=Notes!...`; net profit feeds Notes capital section. |
| **Sch A–H** (Gupta / Ganeshay) | Working schedules (depreciation, debtors, GST, TDS, etc.); some tie to Notes via references. |

### Linking types

| Type | Example (Gupta) | Purpose |
|------|-----------------|--------|
| **BS ← Notes (fixed cells)** | `C9: =Notes!B13` | Face amount = note closing total. |
| **Notes ← P&L** | `Notes!B10: ='Profit & Loss'!C23` | Current-year profit in capital account. |
| **Capital roll-forward in Notes** | `B13: =B8+B9+B10-B11+B12` | Opening + profit + adjustments − drawings. |
| **Partnership: sum partners** | `BS!C9: =Notes!B14+Notes!C14` | B/C are **partner columns**, not CY/PY. |
| **Balance check** | `F8: =C18-C31` | Sources total − assets total (should be 0). |

### Gupta Agency — BS line → Notes anchor map

Use this when cloning the template for another proprietorship (update amounts in Notes; keep references).

| Balance Sheet line | CY formula | PY formula |
|--------------------|------------|------------|
| Proprietor's Capital | `Notes!B13` | `Notes!C13` |
| LT Borrowings – Secured | `Notes!B27` | `Notes!C27` |
| LT Borrowings – Unsecured | `Notes!B44` | `Notes!C44` |
| ST Borrowings | `Notes!B50` | `Notes!C50` |
| Trade Payables | `Notes!B56` | `Notes!C56` |
| Other Current Liabilities | `Notes!B69` | `Notes!C69` |
| ST Provisions | `Notes!B76` | `Notes!C76` |
| PPE | `Notes!B86` | `Notes!C86` |
| Investments | `Notes!B92` | `Notes!C92` |
| CWIP | `Notes!B98` | `Notes!C98` |
| Inventories | `Notes!B109` | `Notes!C109` |
| Trade Receivables | `Notes!B116` | `Notes!C116` |
| Cash | `Notes!B125` | `Notes!C125` |
| ST Loans & Advances | `Notes!B136` | `Notes!C136` |
| Other Current Assets | `Notes!B143` | `Notes!C143` |

### Shree Ganeshay — partnership nuance

- **Note 1** has columns **Vivek Gupta** and **Durga Gupta**; closing balances in `B14`/`C14` (CY) and `B23`/`C23` (PY).
- Balance Sheet **total** partners' capital: `=Notes!B14+Notes!C14` (CY) — sums partners, not years.
- Profit share in Notes should match **partnership deed** (80:20 in this file).

### Shree Balajee — vertical format

- Comparative columns on Notes are **D/E** (not B/C).
- BS pulls e.g. capital from `Notes!$D$21` / `Notes!$E$21`.
- Same hub idea; only column letters differ — when copying, preserve **absolute refs** on totals (`$D$21`).

---

## Function usage summary (all files)

| Function | Where used | CA intent |
|----------|------------|-----------|
| **SUMIF** | HNS Notes | Automated TB aggregation by Map Code |
| **SUM** | All | Note subtotals, BS/P&L totals |
| **IF / ISNUMBER** | HNS Ratios, some checks | Avoid divide-by-zero; “Not meaningful” ratios |
| **ROUND** | Balajee, Ganeshay | Presentation rounding in capital |
| **TEXT** | HNS share capital note | Narrative with share count from Inputs |

---

## What not to break

1. **Do not** type amounts on Balance Sheet face in Family B — always change **Notes** or schedules.
2. **Do not** rename sheets referenced in formulas (`Notes`, `Profit & Loss`, `TB`, `Inputs`).
3. For HNS, **never** edit FS amounts without tracing to **Map Code** on TB.
4. Keep **Check (L − A)** or explicit check rows at **zero** before sign-off.
