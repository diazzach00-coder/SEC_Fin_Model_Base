# Acceptance run: 2026-10-22 (round 2, after the coordinator's rulings)

Command (warm `.cache/`, after all fixes below):

```
python scripts/regress.py --excel --ua "SEC-Model-Framework regression admin@secmodelframework.dev" \
    OKE EPD ET KMI WMB MPLX TRGP KNTK ENB AAPL CVX JPM TRP XOM
```

Result: **12 of 14 pass.** XOM is a documented failure (reported as KNOWN, see section 4).
**MPLX fails** with 3 cash roll-forward FAILs. These come from MPLX's own filings, and the
ruled fix does not apply to them (section 4). The exit code is 1 because of MPLX.

`regress.py` now asserts the following by default, in addition to the brief's assertions:
- The default interest rate is 0-15%. JPM, a bank, gets a NOTE instead of a failure.
- The default cash tax rate is 0-40%.
- Last-year diluted units are within 0.5x-2x of dei shares outstanding.
- No expense concept has the opposite sign to its XBRL fact.

With `--excel` it also recalculates each workbook in Excel, asserts there are no error
cells, and prints the DCF value per unit.

## 1. Final table

`P/W/F/S` = PASS/WARN/FAIL/SKIP checks. `cov` = coverage warnings (required concept
unmapped in the latest year). `bal` = max |BAL_CHECK| over the forecast years. `s` = warm
build seconds (cold builds take 35-58 s, CVX 54 s). `value/unit` = DCF perpetuity / exit
multiple from Excel, with no price entered (equity weight 55%).

| Ticker | Result | Years | Cur | Stmts | P/W/F/S | cov | bal | s | KB | value/unit |
|---|---|---|---|---|---|---|---|---|---|---|
| OKE | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/93 | 0 | 1.5e-11 | 7.6 | 77 | 169.79 / 158.18 |
| EPD | ok | 10 FY2016-FY2025 | USD | 10/10 | 150/10/0/123 | 0 | 1.5e-11 | 6.8 | 74 | 25.01 / 20.48 |
| ET | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/105 | 0 | 2.9e-11 | 8.9 | 81 | 27.54 / 17.20 |
| KMI | ok | 10 FY2016-FY2025 | USD | 10/10 | 151/9/0/135 | 0 | 1.5e-11 | 8.5 | 78 | 20.35 / 15.29 |
| WMB | ok | 10 FY2016-FY2025 | USD | 10/10 | 160/0/0/70 | 0 | 7.3e-12 | 8.3 | 75 | 39.84 / 38.45 |
| MPLX | **FAIL** | 10 FY2016-FY2025 | USD | 10/10 | 155/2/**3**/142 | 0 | 7.3e-12 | 9.4 | 96 | 130.15 / 73.17 |
| TRGP | ok | 10 FY2016-FY2025 | USD | 10/10 | 159/0/0/99 | 0 | 3.6e-12 | 8.2 | 77 | 37.90 / 69.88 |
| KNTK | ok | 9 FY2017-FY2025 | USD | 9/9 | 143/1/0/90 | 0 | 3.6e-12 | 8.5 | 83 | -1.30 / 94.25 |
| ENB | ok | 10 FY2016-FY2025 | CAD | 10/10 | 147/10/0/100 | 0 | 2.9e-11 | 7.4 | 77 | 71.78 / 67.76 |
| AAPL | ok | 10 FY2016-FY2025 | USD | 10/10 | 155/5/0/88 | 0 | 1.2e-10 | 6.4 | 70 | 260.70 / 111.20 |
| CVX | ok | 10 FY2016-FY2025 | USD | 10/10 | 137/19/0/104 | 1 (Cash) | 5.8e-11 | 7.0 | 72 | 104.80 / 99.70 |
| JPM | ok | 10 FY2016-FY2025 | USD | 10/10 | 121/29/0/90 | 1 (Revenue) | 0.0 | 8.3 | 77 | -199.32 / -199.32 |
| TRP | ok (40-F message, as expected) | - | - | - | - | - | - | 0.0 | - | - |
| XOM | KNOWN (documented) | - | - | - | - | - | - | 0.0 | - | - |

Notes:
- **SKIPs** are mostly subtotal checks ("no indented children"), a known limitation.
- **WARNs** are mostly "CF cash vs BS cash" (restricted cash), which is WARN by design.
- **JPM:** Revenue is UNMAPPED, as expected for a bank. Its default interest rate (20.5%,
  deposit interest over debt) is a NOTE, not a failure. Its DCF value (-199) is
  meaningless because the model is not built for banks. JPM is in the list only to show
  that the build does not assume an industrial income statement.
- **CVX:** `Cash: UNMAPPED in FY2025` is a coverage WARN. CVX builds, balances and has no
  FAIL.
- **KNTK:** FY2016 has no filing with all three statements and is dropped. Its perpetuity
  value is about -1.3 because default capex is 23.5% of revenue at 13.5% growth, which
  leaves terminal UFCF near zero. The exit-multiple value is 94.25.
- **Value per unit:** the defaults are mechanical (3-year revenue CAGR capped at ±15%,
  cash tax rate = (tax - deferred) / pretax). OKE's 169.79 comes from its 15%
  acquisition-driven growth cap and a 1.5% cash tax rate. The sign and units bugs behind
  the earlier $3,121 are fixed (section 2). Analysts are expected to overwrite the
  placeholder asset row and the drivers.

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
| OKE, EPD (magnitude check) | The coordinator noted OKE $3,121 and EPD $183 as about 100x too high. | **Root cause: the sign defect, not units.** UNITS are correct (OKE 629.8 MM, EPD 2,161.8 MM). OKE's negative cost of debt (-5.50%) made WACC 2.89%, almost equal to g = 2.5%, so the terminal value was about 263x UFCF. After the fix: OKE WACC 7.80%, value per unit 169.79 / 158.18; EPD WACC 7.26%, 25.01 / 20.48. | - | regress UNITS and rate assertions |
| KNTK | Cash roll-forward FAIL FY2020-21. FY2021 revenue was Altus's 160.6, not BCP's recast 662.0. | **Fix (ruling 2):** each year has one designated supplier, the newest filing with 12-month IS and CF columns and a BS column for that year (`merge.designated_suppliers`). Its values come first, other filings fill only keys it lacks, and `sources[fy]` is that filing. `MergedRow.value_src` records which filing supplied each value. The literal rule still let Altus-only keys win the mapping, so I extended it: tag and label resolution now run on the supplier's rows before anyone else's. | `merge`, `types`, `mapping`, `pipeline` | `test_designated_supplier_is_newest_filing_covering_all_three_statements`, `test_supplier_filing_values_first_and_lacking_keys_from_others`, `test_mapping_prefers_rows_supplied_by_the_years_source_filing` |
| KNTK | After the supplier rule: API cross-check FAIL for NetIncome and OCF in FY2020 (Altus 81.7 vs companyfacts' newest, BCP's recast -1,155.8). | A year's supplier need not be the newest filing. `fy_values(..., accessions)` now prefers the fact from the year's supplier accession and falls back to the newest one. | `companyfacts`, `pipeline` | `test_api_cross_check_reads_the_fact_from_the_years_supplier_filing` |
| ENB, JPM | After the supplier rule: cash roll FAIL FY2016. | Their supplier filings' cash rows use tags outside the CF cash lists (ENB `...IncludingDiscontinuedOperations`, JPM `CashAndDueFromBanks`). Beginning and ending cash therefore came from another filing on a different basis. **Fix:** both the tag and label passes run on the supplier's rows first. | `mapping` | `test_supplier_label_match_beats_tag_match_in_another_filings_row` |
| MPLX (ruling 3) | - | If CFBeginningCash and CFEndingCash resolve to the same row key, the beginning is re-resolved without that row. If they still coincide, the check is SKIP "beginning and ending cash share one row". This does not change MPLX (section 4). | `mapping`, `checks` | `test_beginning_cash_never_shares_the_ending_cash_row`, `test_cash_roll_forward_skips_when_both_ends_share_one_row` |
| CVX (new) | Sanity: default interest rate 45%. | Total debt was 977 (short-term only). CVX/JPM tag long-term debt as `...IncludingCurrentMaturities`, which was missing from the long-term sum group. Long-term totals now close that group and skip the current-maturities group. A one-row sum keeps its As-Reported link. CVX total debt is 40,758. | `concepts`, `mapping` | `test_total_debt_long_term_including_current_plus_short_term_borrowings` |
| JPM | Sanity: default interest rate -162%. | FY2024+ "Interest expense" is `InterestExpenseOperating`, which was missing, so "Net interest income" was taken with sign -1. **Fix:** add it before the net-interest elements. | `concepts` | `test_interest_expense_operating_before_net_interest_income` |
| TRGP, AAPL | Value/unit 793 and 699. Default operating cost was -1.3% (TRGP) and 12% (AAPL) of revenue. | `OperatingExpenses` / `OperatingCostsAndExpenses` were taken as total costs, but they were one line (TRGP) or opex excluding COGS (AAPL). **Fix:** total costs more than 10% away from revenue - operating income are UNMAPPED, so the Model derives cost from revenue - operating income. After the fix: TRGP 37.90 / 69.88, AAPL 260.70 / 111.20. | `mapping` | `test_total_costs_that_do_not_reconcile_to_revenue_less_operating_income_are_dropped` |
| MPLX | InterestExpense UNMAPPED, so the cost of debt was 0% and the perpetuity value 232. | "Net interest and other financial costs" is `InterestAndDebtExpense`. **Fix:** add it to the tag list. Value per unit is now 130.15 / 73.17. | `concepts` | `test_interest_and_debt_expense_maps_to_interest_expense` |
| XOM (ruling 4) | - | `build(..., cik_override=)`, the app sidebar input "CIK override (advanced)" (0 = none), and `regress.py --cik TICKER=CIK`. | `pipeline`, `edgar`, `app` | `test_cik_override_skips_ticker_resolution`, `test_cik_override_unknown_cik_user_facing`, `test_cik_override_zero_means_none_and_value_is_passed` |

## 3. Manual Excel check (COM, `Visible=False`, `CalculateFull`, closed without saving)

Method: open read-only with `DisplayAlerts=False` (a repaired file shows "[Repaired]" in
its caption) and scan every used cell on every sheet for `#NULL! #DIV/0! #VALUE! #REF!
#NAME? #NUM! #N/A`. Then set `Asset Inputs` row 1 FY2026E revenue to 1.10x its value,
recalculate and read again. Finally type a realistic price into `WACC!J6`.

Both workbooks: no repair, **zero error cells** after all three recalculations.

| FY2026E | OKE before | OKE after | EPD before | EPD after |
|---|---|---|---|---|
| Asset Inputs revenue | 38,673.35 | 42,540.69 | 51,207.40 | 56,328.14 |
| Model EBITDA | 8,729.25 | 9,563.58 (up) | 9,135.07 | 10,012.48 (up) |
| Model Cash | 599.26 | 935.08 (up) | 969.00 | 969.00 (held at min cash) |
| Model Total debt | 32,816.00 | 32,816.00 (input) | 34,395.00 | 34,395.00 |
| Model Revolver | 0.00 | 0.00 | 1,481.81 | 1,417.65 (down: repays) |
| Balance check | 0.00 | 0.00 | 0.00 | 0.00 |
| DCF value/unit, perpetuity | 169.79 | 190.78 (up) | 25.01 | 28.69 (up) |
| DCF value/unit, exit | 158.18 | 178.50 (up) | 20.48 | 23.86 (up) |

With no price, WACC is 7.80% for OKE (pre-tax Kd 5.50%, tax 1.50%) and 7.26% for EPD.

After typing a price, "% over / (under) valued" (= price / value - 1) changed from "enter
price" to a number. These readings are taken after the +10% revenue edit, and the price
also moves WACC to market weights:
- OKE at $95: WACC 8.21%, value/unit 172.62 / 214.88, over/under -45.0% / -55.8%.
- EPD at $32: WACC 7.95%, value/unit 23.19 / 25.78, over/under +38.0% / +24.1%.

## 4. Remaining failures and limitations

**MPLX: the ruled fix does not apply (FAIL is correct detection).** Cash roll-forward
FAILs in FY2017, FY2018 and FY2019. Under the supplier rule, those years come from the
FY2018, FY2019 and FY2020 10-Ks. FY2016 now comes from the FY2017 10-K and passes.

I checked those CF pages. They have **two distinct rows**: "Cash, cash equivalents and
restricted cash at beginning of period" (the bare tag) and "... at end of period"
(tag|label). Both show the **end-of-period** values in every column. For example, the
FY2020 10-K R8 shows 15 / 15 / 85 on both rows while the net change is 0 / -70 / 76. This
is a filer XBRL presentation error.

The ruling's premise (both ends resolve to one row key) does not hold, so the new SKIP
does not trigger. Options for the coordinator:
- (a) Accept MPLX as a documented failure in `KNOWN_FAILURES`.
- (b) Also SKIP when the beginning and ending rows carry identical values in a year whose
  net change is non-zero. This is broader than the ruling and could hide a parse error.
- (c) Check the roll-forward against the prior year's ending cash, which would pass
  FY2018-19.

The model is unaffected: cash comes from the BS and the CF uses the net change.

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
