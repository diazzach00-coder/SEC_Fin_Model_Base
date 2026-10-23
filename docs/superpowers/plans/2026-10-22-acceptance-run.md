# Acceptance run: 2026-10-22/23 (final, after the final-review fix wave)

Command (warm `.cache/`, after all fixes below, including section 5):

```
python scripts/regress.py --excel --ua "SEC-Model-Framework regression admin@secmodelframework.dev" \
    OKE EPD ET KMI WMB MPLX TRGP KNTK ENB AAPL CVX JPM TRP XOM
```

Result: **13 of 14 pass; exit code 0.** XOM is a documented failure, reported as KNOWN
(section 4). Every cash roll-forward on the list passes; the only exception is TRGP
FY2016, which is SKIP because an input is missing.

`regress.py` now asserts the following by default, in addition to the brief's assertions:
- The default interest rate is 0-15%. JPM, a bank, gets a NOTE instead of a failure.
- The default tax rate is 0-40%, and within 10 points of the last-3-year effective rate
  (sum of tax / sum of pretax) for filers whose effective rate is above 10%.
- Distributions and DistributionsToNCI are not the same row in the latest year.
- Receivables and Payables are mapped in the latest year when the balance sheet has a
  receivable / payable row. JPM, a bank, gets a NOTE instead of a failure.
- Last-year diluted units are within 0.5x-2x of dei shares outstanding.
- No expense concept has the opposite sign to its XBRL fact.

With `--excel` it also recalculates each workbook in Excel, asserts there are no error
cells, and prints the DCF value per unit.

## 1. Final table

`P/W/F/S` = PASS/WARN/FAIL/SKIP checks. `cov` = coverage warnings (required concept
unmapped in the latest year). `bal` = max |BAL_CHECK| over the forecast years. `s` = warm
build seconds (cold builds take 35-58 s, CVX 54 s). `value/unit` = DCF perpetuity / exit
multiple from Excel, with no price entered (equity weight 55%); `before` is the run at
e2b4231, ahead of the final-review fixes (section 5).

| Ticker | Result | Years | Cur | Stmts | P/W/F/S | cov | bal | s | KB | value/unit | before |
|---|---|---|---|---|---|---|---|---|---|---|---|
| OKE | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/93 | 0 | 1.5e-11 | 7.7 | 77 | 113.02 / 150.37 | 169.81 / 158.20 |
| EPD | ok | 10 FY2016-FY2025 | USD | 10/10 | 150/10/0/123 | 0 | 1.5e-11 | 6.9 | 75 | 26.01 / 20.67 | 25.02 / 20.49 |
| ET | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/105 | 0 | 2.9e-11 | 8.5 | 82 | 26.95 / 17.04 | 27.55 / 17.20 |
| KMI | ok | 10 FY2016-FY2025 | USD | 10/10 | 151/9/0/135 | 0 | 1.5e-11 | 8.5 | 78 | 17.66 / 14.58 | 20.35 / 15.29 |
| WMB | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/70 | 0 | 1.5e-11 | 8.7 | 76 | 30.46 / 36.88 | 39.85 / 38.45 |
| MPLX | ok | 10 FY2016-FY2025 | USD | 10/10 | 158/2/0/142 | 0 | 0.0 | 9.6 | 96 | 130.13 / 73.16 | 130.16 / 73.17 |
| TRGP | ok | 10 FY2016-FY2025 | USD | 10/10 | 159/0/0/99 | 0 | 7.3e-12 | 8.5 | 77 | 21.22 / 66.18 | 37.90 / 69.89 |
| KNTK | ok | 9 FY2017-FY2025 | USD | 9/9 | 143/1/0/90 | 0 | 1.8e-12 | 9.0 | 84 | -6.71 / 94.08 | -1.29 / 94.28 |
| ENB | ok | 10 FY2016-FY2025 | CAD | 10/10 | 147/10/0/100 | 0 | 2.9e-11 | 7.4 | 77 | 57.15 / 65.66 | 71.79 / 67.77 |
| AAPL | ok | 10 FY2016-FY2025 | USD | 10/10 | 155/5/0/88 | 0 | 1.2e-10 | 6.1 | 70 | 260.74 / 111.22 | 260.74 / 111.22 |
| CVX | ok | 10 FY2016-FY2025 | USD | 10/10 | 137/19/0/104 | 1 (Cash) | 5.8e-11 | 6.7 | 72 | 113.88 / 101.23 | 104.80 / 99.70 |
| JPM | ok | 10 FY2016-FY2025 | USD | 10/10 | 121/29/0/90 | 2 (Revenue, Receivables) | 0.0 | 8.0 | 78 | -199.01 / -199.01 | -199.01 / -199.01 |
| TRP | ok (40-F message, as expected) | - | - | - | - | - | - | 0.0 | - | - | - |
| XOM | KNOWN (documented) | - | - | - | - | - | - | 0.0 | - | - | - |

Notes:
- **SKIPs** are mostly subtotal checks ("no indented children"), a known limitation.
- **WARNs** are mostly "CF cash vs BS cash" (restricted cash), which is WARN by design.
- **JPM:** Revenue is UNMAPPED, as expected for a bank. Receivables is UNMAPPED too
  ("Accrued interest and accounts receivable"): a coverage warning, and a NOTE in the
  working-capital sanity assertion. Its default interest rate (20.5%,
  deposit interest over debt) is a NOTE, not a failure. Its DCF value (-199) is
  meaningless because the model is not built for banks. JPM is in the list only to show
  that the build does not assume an industrial income statement.
- **CVX:** `Cash: UNMAPPED in FY2025` is a coverage WARN. CVX builds, balances and has no
  FAIL.
- **KNTK:** FY2016 has no filing with all three statements and is dropped. Its perpetuity
  value is about -6.71 because default capex is 23.5% of revenue at 13.5% growth, which
  leaves terminal UFCF near zero (the effective tax rate, 5.8% instead of 0.6%, pushes it
  further below). The exit-multiple value is 94.08.
- **Value per unit:** the defaults are mechanical (3-year revenue CAGR capped at ±15%,
  tax rate = 3-year average effective rate tax / pretax, spec 7.2). The final-review fixes
  moved the C-corps most: OKE 169.81 -> 113.02 (tax 1.5% -> 23.7%), KMI, WMB, TRGP and ENB
  likewise; the partnerships barely move (EPD 0.3% -> 0.7%, MPLX 0.1% -> 0.2%). The cash
  rate (tax - deferred) / pretax stays on the Model as "Cash tax rate (info)". Analysts are
  expected to overwrite the placeholder asset row and the drivers.

## 2. Fixes

Each fix has its own commit and a unit test that fails without it. The round-1 fixes are
kept: Uncategorized MenuCategory (OKE), R-file currency (ENB), per-year API tag (TRGP), NI
incl. NCI tag (ENB), restricted-cash net change tag (MPLX), NCI nonredeemable tag and label
regex (ET), pretax "before tax" element (ET), SalesRevenueNet anchor (AAPL).

Round 2:

| Ticker(s) | Symptom | Root cause and fix | Module | Test |
|---|---|---|---|---|
| OKE EPD ET KMI WMB TRGP KNTK ENB | Default interest rate -4% to -6%; tax rate clamped to 0; NI to common = NI + NCI; WACC pre-tax cost of debt negative. OKE: $3,121/unit, WACC 2.89% without a price. | R-pages show negated-label rows with the presentation sign. **Fix:** `normalize_expense_signs()` runs after `resolve()`. It puts TotalCostsAndExpenses, DA, Impairments, InterestExpense, IncomeTax and NIToNCI in expense-positive form. Companyfacts is the sign authority per (tag, year) when the tag is in the concept's own list. Custom tags with no fact are flipped only when negative, and only for the never-negative concepts (tax and NCI are excluded). A label-fallback hit on another us-gaap element is left as reported (KMI "(Gain) loss on divestitures and impairments"). `ConceptValue.sign` records the flip and the Mapping sheet writes `-INDEX(...)` per cell. | `mapping`, `types`, `workbook/mapping_sheet`, `pipeline` | `test_negated_label_expense_flipped_to_xbrl_sign`, `test_tax_benefit_matching_fact_sign_is_kept`, `test_custom_tag_without_fact_flipped_only_for_never_negative_concepts`, `test_label_fallback_on_other_us_gaap_element_left_as_reported`, `test_mapping_sheet_applies_per_year_cell_sign` |
| KMI | InterestExpense UNMAPPED ("Interest, net") | Added `InterestIncomeExpenseNonoperatingNet` and `InterestIncomeExpenseNet`, each with tag-level sign -1 (`Concept.tag_signs`). | `concepts`, `mapping` | `test_interest_net_tag_is_income_positive` |
| OKE, EPD (magnitude check) | The coordinator noted OKE $3,121 and EPD $183 as about 100x too high. | **Root cause: the sign defect, not units.** UNITS are correct (OKE 629.8 MM, EPD 2,161.8 MM). OKE's negative cost of debt (-5.50%) made WACC 2.89%, almost equal to g = 2.5%, so the terminal value was about 263x UFCF. After the fix: OKE WACC 7.80%, value per unit 169.81 / 158.20; EPD WACC 7.26%, 25.02 / 20.49 (final run). | - | regress UNITS and rate assertions |
| KNTK | Cash roll-forward FAIL FY2020-21. FY2021 revenue was Altus's 160.6, not BCP's recast 662.0. | **Fix (ruling 2):** each year has one designated supplier, the newest filing with 12-month IS and CF columns and a BS column for that year (`merge.designated_suppliers`). Its values come first, other filings fill only keys it lacks, and `sources[fy]` is that filing. `MergedRow.value_src` records which filing supplied each value. The literal rule still let Altus-only keys win the mapping, so I extended it: tag and label resolution now run on the supplier's rows before anyone else's. | `merge`, `types`, `mapping`, `pipeline` | `test_designated_supplier_is_newest_filing_covering_all_three_statements`, `test_supplier_filing_values_first_and_lacking_keys_from_others`, `test_mapping_prefers_rows_supplied_by_the_years_source_filing` |
| KNTK | After the supplier rule: API cross-check FAIL for NetIncome and OCF in FY2020 (Altus 81.7 vs companyfacts' newest, BCP's recast -1,155.8). | A year's supplier need not be the newest filing. `fy_values(..., accessions)` now prefers the fact from the year's supplier accession and falls back to the newest one. | `companyfacts`, `pipeline` | `test_api_cross_check_reads_the_fact_from_the_years_supplier_filing` |
| ENB, JPM | After the supplier rule: cash roll FAIL FY2016. | Their supplier filings' cash rows use tags outside the CF cash lists (ENB `...IncludingDiscontinuedOperations`, JPM `CashAndDueFromBanks`). Beginning and ending cash therefore came from another filing on a different basis. **Fix:** both the tag and label passes run on the supplier's rows first. | `mapping` | `test_supplier_label_match_beats_tag_match_in_another_filings_row` |
| MPLX (ruling 3, kept) | - | If CFBeginningCash and CFEndingCash resolve to the same row key, the beginning is re-resolved without that row. If they still coincide, the check is SKIP "beginning and ending cash share one row". This does not change MPLX (section 4). | `mapping`, `checks` | `test_beginning_cash_never_shares_the_ending_cash_row`, `test_cash_roll_forward_skips_when_both_ends_share_one_row` |
| CVX (new) | Sanity: default interest rate 45%. | Total debt was 977 (short-term only). CVX/JPM tag long-term debt as `...IncludingCurrentMaturities`, which was missing from the long-term sum group. Long-term totals now close that group and skip the current-maturities group. A one-row sum keeps its As-Reported link. CVX total debt is 40,758. | `concepts`, `mapping` | `test_total_debt_long_term_including_current_plus_short_term_borrowings` |
| JPM | Sanity: default interest rate -162%. | FY2024+ "Interest expense" is `InterestExpenseOperating`, which was missing, so "Net interest income" was taken with sign -1. **Fix:** add it before the net-interest elements. | `concepts` | `test_interest_expense_operating_before_net_interest_income` |
| TRGP, AAPL | Value/unit 793 and 699. Default operating cost was -1.3% (TRGP) and 12% (AAPL) of revenue. | `OperatingExpenses` / `OperatingCostsAndExpenses` were taken as total costs, but they were one line (TRGP) or opex excluding COGS (AAPL). **Fix:** total costs more than 10% away from revenue - operating income are UNMAPPED, so the Model derives cost from revenue - operating income. After the fix: TRGP 37.90 / 69.89, AAPL 260.74 / 111.22 (final run). | `mapping` | `test_total_costs_that_do_not_reconcile_to_revenue_less_operating_income_are_dropped` |
| MPLX | InterestExpense UNMAPPED, so the cost of debt was 0% and the perpetuity value 232. | "Net interest and other financial costs" is `InterestAndDebtExpense`. **Fix:** add it to the tag list. Value per unit is now 130.16 / 73.17. | `concepts` | `test_interest_and_debt_expense_maps_to_interest_expense` |
| MPLX, then KMI/ENB/CVX/JPM/KNTK (ruling: option c) | MPLX FY2017-FY2019 cash roll FAIL: its FY2018-FY2020 10-Ks render the "beginning" row with end-of-period values. Applied literally (opening = prior year's mapped CFEndingCash), the rule turned FY2017 into a FAIL for KMI, ENB, CVX, JPM and MPLX, and FY2021 for KNTK. In those years the ending cash is on a different basis from the prior year's mapped close: restricted cash added by ASU 2016-18 in the FY2018 10-Ks, or KNTK's recast. | The opening balance is the prior fiscal year's ending cash whenever a prior year is in the run. It is read from the **same row** as this year's ending cash (its prior-year comparative, so on the same basis), falling back to the prior year's mapped CFEndingCash. The reported beginning-cash row is used only for the first year. If the prior close does not tie but the year's own reported opening does, the result is WARN ("cash basis changed between source filings"). FAIL only when neither ties. Tolerance, FAIL level and the same-row guard are unchanged. | `checks` | `test_cash_roll_opening_is_prior_year_ending_cash` (beginning row = end values, passes), `test_cash_roll_genuine_mismatch_in_year_two_still_fails`, `test_cash_roll_basis_change_between_source_filings_is_warn_not_fail`, `test_cash_roll_prior_close_read_from_this_years_ending_cash_row` |
| (review) | A TotalCostsAndExpenses row presented negative was compared, still negative, against revenue - operating income and always dropped, so the sign normalisation never reached it. The same-row SKIP returned before the prior-close path ran. | The reconciliation compares absolute values; the later normalisation flips a kept negated total to positive. The same-row SKIP now applies only when there is no prior close (first year). A pinning test covers the one-component sum (status OK, row key and tag kept). `TotalCostsAndExpenses` was added to `regress.py`'s sign assertions. Cached builds of all 12 buildable tickers show identical mapping values and check results before and after this change. | `mapping`, `checks`, `scripts/regress.py` | `test_negated_total_costs_that_reconcile_are_kept_and_flipped_positive`, `test_shared_row_skip_only_without_prior_close`, `test_total_debt_single_component_sum_keeps_the_row_link` |
| XOM (ruling 4) | - | `build(..., cik_override=)`, the app sidebar input "CIK override (advanced)" (0 = none), and `regress.py --cik TICKER=CIK`. | `pipeline`, `edgar`, `app` | `test_cik_override_skips_ticker_resolution`, `test_cik_override_unknown_cik_user_facing`, `test_cik_override_zero_means_none_and_value_is_passed` |

## 3. Manual Excel check (COM, `Visible=False`, `CalculateFull`, closed without saving)

Method: open with `DisplayAlerts=False` (a repaired file shows "[Repaired]" in its
caption) and scan every used cell on every sheet for `#NULL! #DIV/0! #VALUE! #REF! #NAME?
#NUM! #N/A`. Then set `Asset Inputs` row 1 FY2026E revenue to 1.10x its value,
recalculate and read again. Finally type a realistic price into `WACC!J6`.

Both workbooks (regression output after the final-review fixes): no repair, **zero error
cells** after all three recalculations.

| FY2026E | OKE before | OKE after | EPD before | EPD after |
|---|---|---|---|---|
| Asset Inputs revenue | 38,673.35 | 42,540.69 | 51,207.40 | 56,328.14 |
| Model EBITDA | 8,729.25 | 9,563.58 (up) | 9,135.07 | 10,012.48 (up) |
| Model Cash | 78.00 | 78.00 (held at min cash) | 969.00 | 969.00 (held at min cash) |
| Model Total debt | 32,816.00 | 32,816.00 (input) | 34,395.00 | 34,395.00 |
| Model Revolver | 673.54 | 523.08 (down: repays) | 1,367.77 | 1,819.37 (up, see below) |
| Balance check | 0.00 | 0.00 | 0.00 | 0.00 |
| DCF value/unit, perpetuity | 113.02 | 127.17 (up) | 26.01 | 29.73 (up) |
| DCF value/unit, exit | 150.37 | 169.71 (up) | 20.67 | 24.01 (up) |

EPD's revolver now rises with the revenue edit. Before the fixes its working capital was
unmapped (AR and AP days 0), so extra revenue was pure cash. Now AR is 45 days of revenue,
inventory 33 days and AP 10 days of cost, so the +5,121 revenue ties up 890 of working
capital and 433 of capex against +877 EBITDA: year-1 cash from operations falls 19 and
the revolver draws 452 more (the Python model gives the same 1,367.8 -> 1,819.4). The DCF
value still rises because later years keep the higher run-rate.

With no price, WACC is 7.25% for both OKE and EPD (before the fixes: 7.80% and 7.26%; the
after-tax cost of debt now uses the effective tax rate). Drivers year 1 is now a formula
over the Model; after `CalculateFull` every year-1 driver and the MIN_CASH scalar equal
`forecast.driver_defaults` / `scalar_defaults` (OKE max |diff| 1.4e-17, EPD 0.0). The
valuation date is `TODAY()`, so values move slightly from day to day.

After typing a price, "% over / (under) valued" (= price / value - 1) changed from "enter
price" to a number. These readings are taken after the +10% revenue edit, and the price
also moves WACC to market weights:
- OKE at $95: WACC 7.78%, value/unit 108.68 / 205.88, over/under -12.6% / -53.9%.
- EPD at $32: WACC 7.94%, value/unit 24.07 / 25.93, over/under +32.9% / +23.4%.

## 4. Remaining failures and limitations

**MPLX: resolved (ruling: option c).** The FY2018-FY2020 10-Ks render the "beginning"
cash row with end-of-period values. That is a filer XBRL presentation error, verified in
the raw R-pages. The roll-forward now opens from the prior year's close, so the broken row
is not used and all ten years pass.

**XOM: documented limitation.** The ticker now resolves to ExxonMobil Holdings Corp
(CIK 2115436). It became the successor issuer on 2026-07-01 (Form 8-K12B) and has filed no
10-K yet. The history is under Exxon Mobil Corp (CIK 34088). `regress.py` reports XOM as
KNOWN while the failure reads "No 10-K filings found for XOM". It turns into a failure
once XOM starts building, so the documentation cannot go stale silently.
`--cik XOM=34088` (or the app's CIK override) builds FY2016-FY2025 with 0 FAIL and value
per unit 241.65 / 150.74 (147/11/0/90 checks, 1 coverage warning: DilutedUnits). Once the holding company files its FY2026 10-K, a build without
the override will hit TOO_FEW_YEARS.

**KNTK: predecessor break.** FY2017-FY2020 are Altus Midstream as originally reported.
FY2021 onward are BCP, the accounting predecessor recast in the 2022 merger. Each year is
internally consistent, but growth rates across FY2020 to FY2021 are not meaningful.

**Supplier rule trade-off.** A year's three statements now come from one filing. A
restatement that appears only in a later filing's IS/CF comparatives (when that filing's
BS no longer covers the year) is not used. The year comes from the filing that has its
balance sheet, and the API cross-check compares against that same filing's facts.

**Other:**
- Subtotal checks are nearly all SKIP.
- ET counts only nonredeemable NCI.
- AAPL has no interest-expense line (NOTE).
- Default revenue growth hits the ±15% cap for acquisitive names (OKE).
- Banks: the model is not built for them (JPM).

## 5. Final-review fixes (2026-10-23)

A whole-branch review found one Critical and five Important items and two must-fix minors.
Each has its own commit and a test that fails without it.

| Item | Symptom | Fix | Commit |
|---|---|---|---|
| C1 tax rate | TAX_RATE default was the cash rate (tax - deferred) / pretax: OKE 1.5%, KMI 0.9%, WMB 2.9%, TRGP 0.7%, so the forecast and the DCF were nearly untaxed. | Spec 7.2: `TAX_RATE_HIST = MAX(0, MIN(0.35, TAX / PRETAX))`, the driver averages 3 years. A visible RATIO line "Cash tax rate (info)" keeps the cash rate, fed by the hidden DEFTAX helper. OKE 23.7%, KMI 21.1%, WMB 22.6%, TRGP 19.9%, ENB 22.1%. | 4b94c12 |
| I1 TRGP distributions | Common distributions FY2021-FY2025 were the NCI row ("Distributions to noncontrolling interests", FY2025 -33.5 instead of -818.3); DistributionsToNCI took "Contributions from noncontrolling interests" (an inflow), KNTK's took "Net income including noncontrolling interest". | Distributions label excludes noncontrolling/minority and admits "dividends paid"; DistributionsToNCI's label must name a payout; a post-step UNMAPs whichever concept does not own a shared row. | 828a191 |
| I2 working capital | EPD and ET (FY2023/24 on) and ENB (FY2025) had Receivables and Payables UNMAPPED, so the forecast had no working capital; ENB's earlier years were its affiliate rows (90 / 22). | A listed tag with no consolidated row resolves as the DERIVED sum of its member rows (one declared axis, money rows only). Receivables adds `AccountsAndOtherReceivablesNetCurrent` and CVX's `AccountsNotesAndLoansReceivableNetCurrent`; Payables adds `AccountsPayableAndOtherAccruedLiabilitiesCurrent`, `AccountsPayableAndAccruedLiabilitiesCurrent`; labels exclude affiliate/related. Coverage WARN when either is UNMAPPED while the BS shows such a row. EPD AR 6,495 / AP 1,238, ENB 7,081 / 7,555, CVX AR 18,075. | ec40024 |
| I3 Drivers year 1 | Year-1 driver defaults were pasted numbers. | `to_excel(default, ctx)` on the Model's last history column, sheet-qualified: `=AVERAGE(Model!$K$36:$M$36)`, `=Model!M34`, `=N(Model!M80)` for MIN_CASH; constants stay values; input style kept. COM: equal to `forecast.driver_defaults` (max |diff| 1.4e-17). | 9fda876 |
| I4 unmapped pretax | With PretaxIncome unmapped OTHER_INC became -EBIT + interest and the forecast lost its pretax income (toy: FY2026E NI 59.2 instead of 328.3). | OTHER_INC uses `IF(PRETAX_RPT=0, NI_RPT + TAX, PRETAX_RPT)`; PretaxIncome is REQUIRED (coverage warning). No ticker on the list hits it. | 9deef69 |
| I5 equity investments | KMI, WMB (FY2021+) and ENB had EquityInvestments UNMAPPED. | `LongTermInvestments`, `Investments` and the label "^investments$" are accepted only in years where EquityEarnings is mapped (`Concept.gate`). KMI 7,532, WMB 4,559, ENB 21,264 (FY2025). | 2b07b17 |
| I6 toggles | JV_TOGGLE, WC_HOLD and the DCF mid-year cell accepted any value. | List validation "1,0", no blanks. | 62433b0 |
| Minors | README said the latest filing wins across comparatives and pointed at an untracked notebook; pandas unpinned. | README describes the supplier rule; `pandas>=2.1`. | cc361e3 |
| regress | - | Tax-vs-effective-rate, distinct distribution rows and mapped working capital assertions. On the pre-fix mappings they flag OKE, TRGP, EPD and ENB; on the final builds nothing. | d4cc271 |

Mapping changes on the list beyond those above: none. A per-ticker diff of every concept,
year, driver default, coverage warning and check status before and after shows only the
rows named here, and no check changed status.
