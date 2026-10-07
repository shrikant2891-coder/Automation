# CA drafting SOP — balance sheet & financial statements

Use this as a repeatable engagement checklist when preparing the next FY (e.g. **2026–27**) using the templates on branch `Balance-Sheet`.

---

## 1. Engagment setup

| Step | Action |
|------|--------|
| 1.1 | Confirm **entity type**: proprietorship / partnership / private limited / LLP. |
| 1.2 | Confirm **framework**: non-corporate (vertical/horizontal) vs **Schedule III** (companies). |
| 1.3 | Copy the closest prior-year file; rename `FY2025-26` → new year. |
| 1.4 | Roll **opening balances** = prior year closing (audit trail in Notes capital section or TB PY column). |
| 1.5 | Update **Inputs** (company) or header blocks (name, PAN, GSTIN, address, deed reference). |

**Template picker**

- **Pvt Ltd + Tally TB** → `HNSPVTL_FS_FY2025-26_Final.xlsx`
- **Proprietorship with GST/TDS/borrowing schedules** → `Gupta_Agency_...`
- **Partnership with partner columns** → `Shree_Ganeshay_...`
- **Simple vertical proprietorship** → `Shree_Balajee_...`

---

## 2. Trial balance and mapping (mandatory)

### Companies (Family A)

1. Export **trial balance** as at year-end from Tally (ledger, group, Dr, Cr).
2. Paste into **TB** sheet; ensure **double-entry tallies** (total Dr = total Cr).
3. Assign **Map Code** to every ledger (see [HNS-MAP-CODES.md](./HNS-MAP-CODES.md)).
4. New ledger → new row + code; extend Notes `SUMIF` only if code is new to the dictionary.
5. Reconcile **TB net** to **Notes** line-by-line for material items.

### Proprietorship / partnership (Family B)

1. Either maintain TB outside Excel or build totals in **Notes** from schedules.
2. Schedules (Sch A–H) support **tax, GST, depreciation, debtors/creditors** — tie schedule totals to Note lines referenced on BS.
3. **Capital**: P&L net profit must flow to Notes before BS capital line (verify formula links).

---

## 3. Balance sheet — substance (Ind AS / AS / Schedule III)

Work **sources of funds** and **application of funds** (or Schedule III equivalents):

### Liabilities & capital

- [ ] Capital / share capital reconciled to register + deed / SH-6
- [ ] Reserves & surplus / retained earnings = opening ± profit ± dividends ± adjustments
- [ ] **Borrowings** classified current vs non-current (repayment terms, defaults, sub-classification secured/unsecured)
- [ ] **Trade payables** — MSME ageing and **total outstanding dues of micro/small** vs others (Companies Act / UDIN disclosures)
- [ ] **Other current liabilities** — statutory dues, advances from customers, related parties
- [ ] **Provisions** — tax, gratuity, warranty (recognition criteria)

### Assets

- [ ] **PPE / CWIP** — Sch A depreciation, capitalisation policy, CWIP not depreciated
- [ ] **Investments** — classification, cost vs fair value
- [ ] **Inventories** — valuation (cost/NRV), physical verification
- [ ] **Trade receivables** — ageing, ECL / provision for doubtful debts
- [ ] **Cash** — bank confirmations, cash certificate if material
- [ ] **Loans & advances** — related party, recoverability

### Mechanical tie

- [ ] **Total liabilities & capital = total assets** (check cell = 0)
- [ ] Each BS note number ties to **Notes** and supporting schedule

---

## 4. Profit & loss

- [ ] Revenue recognition matches books / GST returns (Sch F reco where used)
- [ ] Expenses complete from TB or Notes build-up
- [ ] **Depreciation** matches Sch A and PPE note
- [ ] **Finance costs** match loan schedules (Sch B/D)
- [ ] **Tax** — current + deferred; rate reconciliation in Sch E (firms)
- [ ] **Partner remuneration** (firms) per deed and s. 40(b) limits where applicable

---

## 5. Tax and regulatory schedules (from your templates)

| Schedule | Typical purpose |
|----------|-----------------|
| Sch A – Depreciation | IT / books depreciation, 15/WDV block |
| Sch B – Borrowings / Debtors | Lender-wise / debtor ageing |
| Sch C – Creditors / GST | MSME, purchase GST reco |
| Sch D – TDS / Loans | TDS deposited; unsecured loans |
| Sch E – Income tax | Provision, advance tax, MAT if any |
| Sch G – ESI | Contribution reco |
| Sch H – Partner remuneration | Remuneration vs deed |

---

## 6. Company-only additions (HNS-style)

- [ ] **Cash flow statement** (indirect) agrees to cash movement
- [ ] **Notes 1–18** (or full Schedule III set) complete
- [ ] **Analytical ratios** — numerator/denominator traceable; document “Not meaningful” where appropriate
- [ ] **Related party** and **contingent liabilities** (Queries sheet → formal note)
- [ ] **EPS** if applicable (AS 20)

---

## 7. Review & sign-off (professional standards)

| Item | Status |
|------|--------|
| Comparative figures agree to **signed prior-year FS** | |
| **Accounting policies** updated for new standards | |
| **Going concern** assessed | |
| **UDIN** generated and filled in Inputs / signature block | |
| **CARO 2020** / **internal financial controls** (companies) | |
| Partner review of **materiality** and **key audit matters** | |
| Client authorisation for **signatory** names on FS | |

---

## 8. Year-end roll-forward (file maintenance)

1. Save as new file `..._FY2026-27.xlsx`.
2. Copy **closing column → opening** in Notes capital and movement schedules.
3. Shift **comparative column** (PY becomes earlier year; add new CY column if needed).
4. Clear **UDIN** and signing date until final audit.
5. Run `python3 scripts/validate_balance_sheets.py` and fix any non-zero BS difference.

---

## 9. How I can assist on future drafts (as CA workflow)

When you start a new client year, provide:

1. **Entity type** and which template to clone  
2. **Trial balance** export (for companies) or schedule totals (for firms)  
3. **Changes**: new loans, partners, GST issues, related parties  
4. **Prior signed FS** PDF for comparative tie  

I can then: map ledgers to codes, trace Notes ↔ BS links, flag check failures, and draft disclosure text in Notes consistent with your existing formats.
