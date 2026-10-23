# Acceptance run: 2026-10-22

Command (warm `.cache/`, after all fixes below):

```
python scripts/regress.py --ua "SEC-Model-Framework regression admin@secmodelframework.dev" \
    OKE EPD ET KMI WMB MPLX TRGP KNTK ENB AAPL XOM JPM TRP
```

Result: **10 of 13 pass. MPLX, KNTK and XOM fail.** None of the three is fixed by a
data-level change. MPLX's failure comes from the filer's own data. KNTK and XOM each need a
design decision (sections 4 and 5).

The regression assertions pass for every midstream name except MPLX and KNTK. Even so, the
models built for OKE, EPD, ET, KMI, WMB, TRGP, KNTK and ENB **still have sign errors that
change their outputs** (section 3). The acceptance assertions do not check for these;
`regress.py` reports them as `SANITY` lines.

## 1. Final table

`P/W/F/S` = PASS/WARN/FAIL/SKIP check counts. `cov` = coverage warnings (required concept
unmapped in the latest year). `bal` = max |BAL_CHECK| over the forecast years. `warm s` =
time for this run. `cold s` = time for the first run (empty cache, before fixes).

| Ticker | Result | Years | Cur | Stmts | P/W/F/S | cov | bal | warm s | cold s | KB | SANITY |
|---|---|---|---|---|---|---|---|---|---|---|---|
| OKE | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/93 | 0 | 1.5e-11 | 7.1 | 9.2* | 77 | 3 |
| EPD | ok | 10 FY2016-FY2025 | USD | 10/10 | 150/10/0/123 | 0 | 0.0 | 6.5 | 52.0 | 74 | 4 |
| ET | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/105 | 0 | 2.9e-11 | 8.2 | 53.4 | 81 | 2 |
| KMI | ok | 10 FY2016-FY2025 | USD | 10/10 | 150/10/0/135 | 0 | 2.9e-11 | 8.2 | 57.8 | 78 | 3 |
| WMB | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/70 | 0 | 1.5e-11 | 8.1 | 41.9 | 75 | 2 |
| MPLX | **FAIL** | 10 FY2016-FY2025 | USD | 10/10 | 154/3/**3**/142 | 0 | 7.3e-12 | 9.5 | 44.1 | 96 | 1 |
| TRGP | ok | 10 FY2016-FY2025 | USD | 10/10 | 159/0/0/99 | 0 | 7.3e-12 | 8.3 | 46.3 | 77 | 2 |
| KNTK | **FAIL** | 9 FY2017-FY2025 | USD | 9/9 | 137/3/**2**/89 | 0 | 3.6e-12 | 8.8 | 35.1 | 83 | 1 |
| ENB | ok | 10 FY2016-FY2025 | CAD | 10/10 | 148/10/0/99 | 0 | 0.0 | 8.2 | 40.8 | 77 | 3 |
| AAPL | ok | 10 FY2016-FY2025 | USD | 10/10 | 155/5/0/88 | 0 | 1.2e-10 | 7.0 | 35.9 | 70 | 1 |
| XOM | **FAIL** | - | - | - | - | - | - | 0.0 | 0.7 | - | - |
| JPM | ok | 10 FY2016-FY2025 | USD | 10/10 | 120/30/0/90 | 1 | 9.3e-10 | 8.8 | 40.0 | 77 | 0 |
| TRP | ok (expected 40-F error) | - | - | - | - | - | - | 0.0 | 0.8 | - | - |

\* OKE's cache was already warm from Task 13. Cold builds take 35-58 s per ticker, about
8.5 minutes for the whole list. ENB's cold time is from its first run, which stopped at
the currency error before the build finished. XOM fails with "No 10-K filings found for XOM." (section 4).

Notes on the table:
- Most SKIPs are subtotal checks ("no indented children"). This is a known limitation.
- Most WARNs are "CF cash vs BS cash". These are restricted cash differences and the
  check is WARN by design.
- JPM: `Revenue: UNMAPPED in FY2025` is expected for a bank. It still has zero FAIL.
- TRP returns "TRP files Form 40-F (no US-GAAP XBRL statements); not supported." This is
  the expected result and the assertion requires it.
- ENB: the currency is CAD. ENB's 10-K/As carry no statements (FilingSummary lists only
  the cover), so the original 10-K is used each year. This is correct (spec §5.1). The
  task context expected the amendment to be chosen, but that would only apply to an
  amendment that contains statements.
- KNTK drops FY2016. Its earliest annual filings have no full statement set.

## 2. Fixes (one commit per root cause, each with a unit test that fails without it)

| Ticker | Symptom | Root cause | Module | Test |
|---|---|---|---|---|
| OKE | FY2021 IS not found in its own 10-K. The year was covered only by comparatives. | `pick_statement_reports` admitted only `MenuCategory == "Statements"`, and OKE's FY2021 IS is listed under `Uncategorized`. Fix: admit `Uncategorized`, ranked after `Statements` at equal score. The anchor-tag check still confirms the pick. | `edgar` | `test_pick_admits_uncategorized_statement`, `test_pick_statements_category_before_uncategorized_at_equal_score` |
| ENB | `UserFacingError: Currency changes across filings (USD→CAD)` | The R-file currency regex read only `XXX ($)` in the title. ENB's FY2017 titles read `... - CAD CAD in Millions`, and that also left the money scale at 1e-6. Several FY2018/FY2022-24 IS/CF pages put the currency only in the column headers (`Dec. 31, 2022 CAD ($)`). Fix: use the currency in a kept column header first, then `XXX ($)` or `XXX in Millions` in the title. The money scale now accepts a currency code in place of `$`. | `rfile` | `test_currency_and_scale_from_code_in_millions_title`, `test_currency_from_column_header_when_title_has_only_dollar_scale` |
| TRGP | `API cross-check: Revenue FY2016` FAIL (6,690.9 vs 5,586.7) | `pipeline._facts_values` read every year's companyfacts value with the tag mapped in the **latest** year (`Revenues`). FY2016 is mapped via `RevenueFromContractWithCustomer...`, whose companyfacts value is 6,690.9. Fix: read each year with its own mapped tag. | `pipeline` | `test_api_cross_check_reads_each_year_with_that_years_tag` |
| ENB | `NI ties IS=CF` FAIL and NetIncome API FAIL, FY2016-FY2019 | Consolidated "Earnings" is tagged `IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest` and there is no `ProfitLoss` row. NetIncome therefore fell to `NetIncomeLoss`, which excludes NCI. Fix: add the tag between `ProfitLoss` and `NetIncomeLoss` (NetIncome and CFNetIncome). | `concepts` | `test_net_income_prefers_continuing_ops_incl_nci_over_controlling_interest` |
| MPLX | `Cash roll-forward` FAIL FY2016-17. FY2018+ net change unmapped. | MPLX tags the restricted-cash net change with `...PeriodIncreaseDecreaseExcludingExchangeRateEffect`, which was not in the tag list. FY2016-17 therefore fell back to an older filing's cash-only net change, which is on a different basis. Fix: add the tag. Three FAILs remain from the filer's own data (section 4). | `concepts` | `test_net_change_in_cash_restricted_basis_excluding_fx_tag` |
| ET | FY2025 NI to NCI = 6,058 (its pretax row). NI to common was -350. | `NetIncomeLossAttributableToNoncontrollingInterest` was replaced by the `...Nonredeemable...` element. The bare `noncontrolling` label regex then matched "Income (Loss), Including Portion Attributable to Noncontrolling Interest". Fix: add the tag, and make the label fallback skip "including portion" rows. Found in the post-run mapping audit; no check caught it. | `concepts` | `test_nci_income_nonredeemable_tag_and_label_skips_including_portion_total` |
| ET | PretaxIncome UNMAPPED FY2022+. The Model's history pretax line read 0, and forecast NI collapsed because the "Discontinued ops & other" plug absorbed NI + tax and then went to 0. | ET uses the newer element `IncomeLossIncludingPortionAttributableToNoncontrollingInterest` ("..., before Tax"). Fix: add it to PretaxIncome. Found in the post-run mapping audit. | `concepts` | `test_pretax_income_newer_before_tax_element` |
| AAPL | FY2016 dropped (9 years). FY2016/FY2017 10-K IS not located. | The pre-ASC 606 IS uses `SalesRevenueNet`, which was not an IS anchor tag. Fix: add it to the anchors. | `edgar` | `test_locate_statements_accepts_pre_asc606_sales_revenue_net_anchor` |

`scripts/regress.py` (new) implements the brief's contract: `--ua`, `--years`, `--out`, one
line per ticker, a final table, collected assertions and exit code 1. It also adds
`--min-years` (default 5), `-v` (echo pipeline progress) and the `SANITY` notes (`--strict`
turns them into assertions).

## 3. BLOCKED (design): negated-label signs are carried into the model

**Evidence.** R-pages show values with the filer's *presentation* sign. When a filer uses
a negated label ("Interest expense (1,783)"), the R-page shows the XBRL value negated. The
mapping copies the displayed sign (`Concept.sign` is a fixed per-concept constant). The
model spec assumes the opposite convention: expenses are positive. Its INTEREST, TAX and
NCI lines use the mapped values directly. The comparison against companyfacts below uses
the same tag and year, FY2025:

| Ticker | InterestExpense | IncomeTax | NIToNCI |
|---|---|---|---|
| OKE | -1,783 vs XBRL +1,783 | -1,028 vs +1,028 | |
| EPD | -1,401 vs +1,401 | -23 vs +23 | -62 vs +62 |
| ET | -3,474 vs +3,474 | | |
| KMI | UNMAPPED ("Interest, net" = `InterestIncomeExpenseNonoperatingNet`) | -832 vs +832 | -104 vs +104 |
| WMB | -1,442 vs +1,442 | | |
| TRGP | (negative default rate) | -529.7 vs +529.7 | |
| KNTK | (negative default rate) | | |
| ENB | -5,023 vs +5,023 | -2,004 vs +2,004 | |

**Consequences** (seen in the Excel check, section 6):
- The default interest rate on debt is negative (OKE -5.50%, EPD -4.23%, ET -5.42%,
  WMB -5.13%, TRGP -5.60%, KNTK -6.40%, ENB -4.86%). Forecast interest is therefore
  *income*.
- The WACC sheet's pre-tax cost of debt is negative. With no price entered, OKE's WACC is
  2.89% and the perpetuity value per unit is $3,121.
- A negative tax makes the cash tax rate clamp to 0%. A negative NI to NCI makes NI to
  common = NI + NCI.
- Historical net income still ties, because the Model's "Other income" plug absorbs the
  error. The balance check is unaffected, so the regression assertions cannot see this.

**Proposed fix** (a mapping-rule change, which is why it is not made here):
1. Mapping values follow the XBRL sign. Every concept's us-gaap tags already use the
   model's sign convention.
2. For each (concept, year) with a us-gaap tag: if the companyfacts value for that tag
   and year equals the negative of the R-page value (within tolerance), negate it. Record
   the flip on `ConceptValue` (a new defaulted field). The Mapping sheet writes
   `-INDEX(...)` for that year, the same way `concept.sign == -1` does now, but per year.
3. Custom tags with no facts take the flip decided for the same row key in the nearest
   year that has evidence.
4. KMI-style `InterestIncomeExpenseNonoperatingNet` ("Interest, net") would need a
   per-tag sign of -1. `Concept` currently has only a per-concept sign.

`regress.py --strict` turns the SANITY notes into assertions, so it will verify the fix.

## 4. Remaining failures

**MPLX: data limitation (the FAIL is correct).** Cash roll-forward FAILs in FY2016,
FY2017 and FY2018. In MPLX's FY2018, FY2019 and FY2020 10-K CF R-pages, the "Cash, cash
equivalents and restricted cash at beginning of period" row shows the *ending* balances in
every column. FY2020 10-K R8 shows 15 / 15 / 85 on both rows while the net change is
0 / -70 / 76. The FY2021 10-K is correct. Each FAIL is therefore a correct detection of an
inconsistency in the source filing. The model does not use the beginning-cash row: cash
comes from the BS, and the CF uses the net change. The workbook is usable. A roll-forward
based on prior-year ending cash would pass FY2017-18 but not FY2016 (no prior year in the
build). I did not change the check's definition.

**KNTK: BLOCKED (merge/mapping rule).** Kinetik was formed in February 2022 when Altus
Midstream (the registrant, CIK 1692787) merged with BCP. BCP is the accounting predecessor,
so the FY2022+ 10-Ks recast 2020-2021 to BCP's numbers. The (year, key) merge keeps rows
that only older filings contain. The mapping then prefers those rows by tag order or
merged row order, so a single year mixes the two entities:
- FY2020 Revenue is Altus's 148.4 (`us-gaap_Revenues`, from Altus's own 10-K). The
  FY2022 10-K reports 410.2 under `RevenueFromContractWithCustomer...`.
- FY2021 CF ending cash is Altus's 131.963, while BS cash is BCP's 18.729.
- FY2021 NetIncome and OperatingIncome are unmapped.

The two cash roll-forward FAILs (FY2020, FY2021) are this mix. Before the TRGP fix, the
Revenue API cross-check also failed for FY2020-21. After it, each year is compared with
the companyfacts value for the same tag, which is Altus's figure, so those years pass.
The mix is still caught by the cash roll FAILs. **Proposed fix:** record the supplying
filing per (row, year) in the merge (`MergedRow`). For each year, the mapping first
resolves against rows supplied by that year's primary source filing (`sources[fy]`), and
falls back to older-only rows only when that finds nothing.

**XOM: BLOCKED (ticker resolution).** `company_tickers.json` now maps XOM to
**ExxonMobil Holdings Corp, CIK 2115436**. It became the successor issuer on 2026-07-01
(Form 8-K12B) and has filed no 10-K yet. The history is under the predecessor **Exxon
Mobil Corp, CIK 34088**. That predecessor co-registered the 2026-06-30 10-Q (accession
0000034088-26-000093), which is listed under both CIKs. A scratch build with XOM pointed
at CIK 34088 gives FY2016-FY2025, 148 PASS / 11 WARN / **0 FAIL**, bal 5.8e-11. The data
builds; only the ticker resolution fails. **Proposed fix:** when a CIK has no annual
filings but has an 8-K12B, find the predecessor through a co-registered accession. Use
its annual filings and companyfacts, and also the successor's own 10-Ks once they exist.
Alternatively, accept a numeric CIK in the ticker box as a manual workaround. Once the
holding company files its FY2026 10-K, a successor-only build will hit TOO_FEW_YEARS, so
this needs a decision before then.

## 5. Other known limitations
- Subtotal checks are nearly all SKIP on modern filings (indent detection).
- ET counts only the nonredeemable NCI share in NI to NCI. Redeemable NCI (67 in FY2025)
  is not added.
- InterestExpense is UNMAPPED for KMI, MPLX and AAPL in FY2025, so forecast interest = 0
  for those (see section 3).
- The default revenue growth hits the ±15% cap for acquisitive names (OKE: 15%). This
  follows the spec (3-year CAGR, capped).

## 6. Manual Excel check (COM, Excel `Visible=False`, `CalculateFull`, closed without saving)

Method: open the workbook read-only through `win32com` with `DisplayAlerts=False`. A
repaired workbook would show "[Repaired]" in its window caption. Scan every used cell on
every sheet for error values (`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`,
`#N/A`). Read the first forecast year (FY2026E). Overwrite `Asset Inputs` row 1
(Consolidated) revenue with 1.10x its value, recalculate and read again. Then type 100 into
the WACC share price cell (`WACC!J6`), recalculate and read "% over / (under) valued".

Both workbooks: no repair, **zero error cells** after each of the three recalculations
(including no `#N/A`).

| FY2026E | OKE before | OKE after | EPD before | EPD after |
|---|---|---|---|---|
| Asset Inputs revenue (N6) | 38,673.35 | 42,540.69 | 51,207.40 | 56,328.14 |
| Model EBITDA | 8,729.25 | 9,563.58 (up) | 9,135.07 | 10,012.48 (up) |
| Model Cash | 720.37 | 1,068.68 (up) | 969.00 | 969.00 (held at min cash) |
| Model Total debt | 32,816.00 | 32,816.00 (input-driven) | 34,395.00 | 34,395.00 |
| Model Revolver | 0.00 | 0.00 (no draw needed) | 1,235.77 | 1,168.998 (down: repays) |
| Balance check | 0.00 | 0.00 | 0.00 | 0.00 |
| DCF value per unit, perpetuity | 3,121.55 | 3,423.93 (up) | 183.43 | 201.66 (up) |
| DCF value per unit, exit multiple | 201.48 | 226.02 (up) | 25.86 | 29.74 (up) |

Every figure moved in the expected direction. For OKE, extra cash builds above minimum
cash with no revolver. For EPD, cash is pinned at minimum cash, so the revolver repays
instead. Debt is an input (issuance 0), so it does not move.

"% over / (under) valued" showed "enter price" before a price was entered. After entering
100 it became numeric:
- OKE: -83.6% (perpetuity) and -62.2% (exit). WACC 4.53%; value per unit 609.52 / 264.23.
- EPD: +319.7% and +35.9%. WACC 7.88%; value per unit 23.82 / 73.59.

**The mechanics work, but the absolute valuation numbers are not credible** until the
sign defect in section 3 is fixed. OKE's pre-tax cost of debt is -5.50% and its WACC
without a price is 2.89%.
