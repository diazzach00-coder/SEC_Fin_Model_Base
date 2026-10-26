> Historical design document (2026-10-22). The tool has since been generalised for any US-GAAP 10-K filer;
> see `docs/superpowers/plans/2026-10-25-general-purpose.md` for the current release's decisions.
> Client references below are historical context.

# SEC 3-Statement Model Builder — Design

Date: 2026-10-22
Status: approved 2026-10-22 (JV toggle decision incorporated)
Target: working demo for East Daley on Wednesday 2026-10-28

## 1. Intent

East Daley Analytics builds asset-level forecasts for midstream energy companies
(pipelines, gathering, processing). Their models stop at cash flow and debt; they
have no linked income statement, balance sheet and cash flow statement, and they
do not model working capital. This tool gives them that missing piece.

An analyst types a ticker once. The tool pulls that company's full SEC 10-K
history (income statement, balance sheet, cash flow statement, in reported order,
oldest to newest), checks that it ties, and writes an Excel workbook where:

- the history is as-reported, by XBRL tag, so the same line is the same line in
  every year even when the company renames it;
- a 5-year forecast runs off an **Asset Inputs** block (revenue and operating
  cost per asset per year), which is where East Daley's own asset data will
  land once it exists;
- everything below EBITDA (D&A, interest, tax, capex, distributions, debt,
  working capital) runs off simple yellow drivers that default to the company's
  own history, so the balance sheet balances on first open;
- a WACC build-up and a DCF with both perpetuity-growth and EBITDA-exit-multiple
  terminal values sit on top, following the structure of the existing
  `DCFWacc Templates.xlsx`.

This is a "here is what I have been building" piece. The first two minutes in
front of East Daley have to show: type OKE, get a workbook that ties, change an
asset's revenue, watch EBITDA, cash, debt and the DCF move.

## 2. Goals, non-goals, anti-goals

**Goals**
- Works for any US-listed company that files 10-K with XBRL (midstream first:
  OKE, EPD, ET, KMI, WMB, MPLX, TRGP, KNTK, ENB; but also any other 10-K filer).
- Numbers are the SEC's numbers. Every historical cell traces to a filing,
  statement, tag and period.
- Tie-outs are automatic and visible: assets = liabilities + equity; IS net
  income = CF net income; cash roll-forward closes; key figures agree with the
  SEC companyfacts API.
- All forecast and valuation cells are live Excel formulas. An analyst can
  trace any number back to an input or an SEC figure.
- Partnerships (units, distributions, "net income attributable to partners")
  work as well as C-corps.
- Never crashes on a supported ticker. Missing pieces degrade to a visible
  "UNMAPPED" or "not reported" marker, never a silent zero and never an
  exception.

**Non-goals**
- Quarterly (10-Q) history. The existing `SEC_10Q_Builder.ipynb` stays as-is;
  this project is annual only.
- 40-F / 20-F filers (TC Energy, Pembina). They have no US-GAAP XBRL
  statements. The app tells the user plainly and stops.
- Segment or asset-level *history* from the filings. The asset block is for the
  forecast; history is consolidated.
- Live market data (share price, beta). These are yellow inputs. See §9.
- Replacing East Daley's asset models. The asset block is the hand-off point,
  not a reimplementation.

**Anti-goals (what would make this a failure even if it "works")**
- A balance sheet that does not balance, or a forecast that balances only
  because of a hidden plug.
- Rows that split or duplicate across years because a label changed.
- Hard-coded forecast numbers in cells where a formula should be.
- A model that looks finished but has `#REF!` or Excel data tables that do not
  calculate (the template's sensitivity tables currently show `#REF!`).

## 3. Constraints

- Stack stays as it is: Python, Streamlit, requests, pandas, openpyxl,
  beautifulsoup4. No new runtime dependencies without a reason written in this
  doc. Must deploy on Streamlit Community Cloud.
- SEC fair-access policy: descriptive User-Agent with email, ≤10 requests/sec,
  retries with backoff. Cache every SEC response on disk for the session so a
  re-run of the same ticker costs no network.
- Excel output must open cleanly in Excel desktop with no repair prompt and no
  `#REF!`. Sensitivity grids are built from ordinary formulas, not Excel Data
  Tables (openpyxl cannot write working data tables).
- openpyxl's strengths and limits set the Excel feature set: formulas, number
  formats, fills, defined names, freeze panes, grouping, basic bar charts. No
  macros.

## 4. User flow (Streamlit)

Sidebar: Ticker, SEC User-Agent email, Years of history (default 10, max what
the company has), Forecast years (default 5), Build button.

Main pane, in order:
1. Progress log (resolve ticker, list filings, download statements, parse,
   merge, check, write).
2. **Data checks panel**: one row per check with PASS / WARN / FAIL, the
   numbers compared, and the tolerance. FAIL does not stop the build; it marks
   the workbook cover sheet.
3. Preview tabs: Income Statement, Balance Sheet, Cash Flow, Mapping (which
   as-reported line feeds each model driver).
4. Download button: `{TICKER}_3Stmt_{YYYY-MM-DD}.xlsx`.

If the ticker is not found, or has no 10-K with XBRL statements, the app says
so in one sentence and stops. No traceback is ever shown to the user; the full
traceback goes to the server log.

## 5. Data pipeline

### 5.1 Resolve ticker and filings
- `https://www.sec.gov/files/company_tickers.json` → CIK and name. Also try
  `company_tickers_exchange.json` as a fallback for tickers with a `.`/`-`
  class suffix (BRK-B style).
- `https://data.sec.gov/submissions/CIK##########.json`. Read
  `filings.recent` **and** every file listed under `filings.files` (older
  filings are paginated; large filers overflow `recent`).
- Keep forms `10-K`, `10-K/A`, `10-KT`. Group by fiscal-year end
  (`reportDate`). For each fiscal year keep the **latest** filing that actually
  contains the three statements (see 5.2); if an amendment has no statements
  (common for Part III-only 10-K/As), fall back to the original. ENB files a
  full 10-K/A every year; ET filed one for FY2024.
- Fiscal-year labels come from the fiscal-year-end date (`FY2024` for a
  2024-12-31 end). For non-December year ends label by the year the fiscal year
  ends in, and show the actual end date in the header row.

### 5.2 Find the three statements in each filing
- Fetch `FilingSummary.xml` from the filing folder. It lists every R-file with
  `ShortName`, `LongName`, `MenuCategory`, `HtmlFileName`. This is more
  reliable than `MetaLinks.json` (which the current app uses and which is
  missing on older filings).
- Candidates are reports with `MenuCategory == Statements`. Score names with
  the existing precise/broad keyword lists, and **exclude** parenthetical,
  comprehensive income, equity, and any name containing "Details".
- Prefer the report whose rows contain the anchor tag for that statement
  (`Revenues`/`RevenueFromContractWithCustomer...` or `CostsAndExpenses` for
  IS; `Assets` for BS; `NetCashProvidedByUsedInOperatingActivities` for CF).
  Name scoring picks the candidates; the anchor tag confirms.
- If a statement cannot be found in a filing, record it as missing for that
  filing; the year may still be covered by the next filing's comparative
  columns (10-K IS and CF carry three years, BS carries two).

### 5.3 Parse an R-file
Each R-file is one HTML table. From it take:
- **Header**: the statement title with unit text (`USD ($) $ in Millions`,
  `shares in Thousands`), and the period columns. Period header rows give
  `12 Months Ended` + a date per column, or for the BS a date per column.
  Columns are keyed by their **period end date and duration**, not by
  position. Columns that are not full-year periods (e.g. a 1-month stub after
  a fiscal-year change) are kept but flagged, and excluded from the model.
- **Rows**: the first cell holds the label and an `onclick="...defref_<prefix>_<Tag>..."`
  anchor with the XBRL element name. Prefix is `us-gaap`, `ifrs-full`, or the
  company's own (`epd`, `et`, ...). Indent level comes from the cell's
  `padding-left` / nested `div` style; abstract rows (tag ends in `Abstract`)
  are section headers.
- **Values**: numbers with `$`, commas, parentheses, `—`; footnote markers
  like `[1]` are stripped. Scale: dollar rows use the `$ in ...` scale; share
  and unit rows use the `shares in ...` scale; per-share rows (`EarningsPerShare*`,
  `...PerUnit`, `...PerShare`) are not scaled. Row type comes from the tag's
  unit in the companyfacts API when the tag is standard, otherwise from the
  label. Everything is stored in **$MM**, **MM shares/units**, or per-share.
- Rows with dimensional qualifiers (the same tag appearing more than once in a
  statement, e.g. revenue by class) are keyed by `tag + label` so they stay
  distinct.

### 5.4 Merge across filings
- Row **order and labels** come from the most recent filing. Rows that only
  exist in older filings are inserted after the row that precedes them in
  their own filing (or at the end of their section), with their last-used
  label and a note that they are no longer reported.
- Row **key** is the XBRL tag (plus label for dimensional duplicates). A
  label change with the same tag is one row. A tag change with the same label
  stays two rows; the mapping layer (5.6) is where a human-meaningful concept
  pulls the right one per year.
- For each fiscal year, the **value is taken from the latest filing that
  reports that year**, so restatements win. The cover sheet lists, per year,
  which filing supplied it.
- A per-year "source" record (accession, R-file, scale) is kept for the
  Sources sheet.

### 5.5 Checks (every build, every ticker)
Tolerances are absolute $1MM or 0.5% of the larger figure, whichever is
larger, to absorb rounding.

| Check | Rule | On failure |
|---|---|---|
| BS balances | `Assets` = `LiabilitiesAndStockholdersEquity` (or `LiabilitiesAndPartnersCapital`) each year | FAIL |
| NI ties | IS net income (`ProfitLoss` → `NetIncomeLoss`) = first line of CF operating section each year | FAIL |
| Cash roll | CF beginning cash + net change = CF ending cash each year | FAIL |
| CF cash vs BS cash | CF ending cash vs BS `CashAndCashEquivalentsAtCarryingValue` | WARN only (restricted cash is often included in CF cash) |
| Subtotals | Any row that is a `Total ...` and has indented children above it: children sum to it | WARN |
| API cross-check | Revenue, net income, total assets, operating cash flow vs companyfacts API `fp=FY` values for the same period | WARN if different, FAIL if different by >5% |
| Coverage | Each requested year has IS, BS and CF; forecast anchors (revenue, EBITDA components, cash, debt) are mapped | WARN per missing item |

Results are shown in the app and written to the **Checks** sheet.

### 5.6 Concept mapping (history → model drivers)
A fixed list of model concepts, each with an ordered list of accepted XBRL
tags, and a label-regex fallback for custom tags. Example:

| Concept | Tags (in order) | Label fallback |
|---|---|---|
| Revenue | `Revenues`, `RevenueFromContractWithCustomerExcludingAssessedTax`, `SalesRevenueNet` | `^total revenues?` |
| D&A | `DepreciationDepletionAndAmortization`, `DepreciationAndAmortization`, `DepreciationAmortizationAndAccretionNet`, `Depreciation` (+ `AmortizationOfIntangibleAssets`) | `depreciation` |
| Interest expense | `InterestExpense`, `InterestExpenseNonoperating`, `InterestExpenseDebt` | `interest expense` |
| Income tax | `IncomeTaxExpenseBenefit` | `income tax` |
| Net income (consolidated) | `ProfitLoss`, `NetIncomeLoss` | — |
| NI to noncontrolling | `NetIncomeLossAttributableToNoncontrollingInterest` | `noncontrolling` |
| Operating cash flow | `NetCashProvidedByUsedInOperatingActivities` | — |
| Capex | `PaymentsToAcquirePropertyPlantAndEquipment`, `PaymentsToAcquireProductiveAssets`, `PaymentsForCapitalImprovements` | `capital expenditures` |
| Distributions / dividends | `PaymentsOfDividendsCommonStock`, `PaymentsOfDividends`, `DistributionMadeToLimitedPartnerCashDistributionsPaid`, `DistributionMadeToMemberOrLimitedPartnerCashDistributionsPaid`, `PaymentsOfCapitalDistribution` | `distributions? (paid\|to)` |
| Cash | `CashAndCashEquivalentsAtCarryingValue`, `CashCashEquivalentsRestrictedCash...` | — |
| Total debt | sum of `LongTermDebtNoncurrent` + `LongTermDebtCurrent` (+ `DebtCurrent`, `CommercialPaper`, `ShortTermBorrowings`) or `LongTermDebt` | `debt` |
| Equity | `StockholdersEquity`, `PartnersCapital`, `StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest` | — |
| Diluted shares/units | `WeightedAverageNumberOfDilutedSharesOutstanding`, `WeightedAverageLimitedPartnershipUnitsOutstandingDiluted` | — |
| Shares/units outstanding | `dei:EntityCommonStockSharesOutstanding`, `CommonStockSharesOutstanding`, `LimitedPartnersCapitalAccountUnitsOutstanding` | — |
| Working capital items | receivables, inventory, payables, accrued (standard tags) | labels |
| Operating income | `OperatingIncomeLoss` | `operating income` |
| Total costs and expenses | `CostsAndExpenses`, `OperatingExpenses`, `OperatingCostsAndExpenses` | `total costs and expenses` |
| Impairments | `AssetImpairmentCharges`, `OtherAssetImpairmentCharges`, `GoodwillImpairmentLoss`, `ImpairmentOfLongLivedAssetsHeldForUse` | `impairment` |
| Gains on sale | `GainLossOnDispositionOfAssets`, `GainLossOnDispositionOfAssets1`, `GainLossOnSaleOfBusiness` | `gain.*(sale\|disposition)` |
| Equity earnings | `IncomeLossFromEquityMethodInvestments` | `equity (in )?(net )?(earnings\|income)` |
| Distributions from affiliates | `EquityMethodInvestmentDividendsOrDistributions` | `distributions received` |
| Distributions to NCI | `PaymentsOfDividendsMinorityInterest`, `PaymentsOfDistributionsToAffiliates` | `noncontrolling` |
| Net PP&E | `PropertyPlantAndEquipmentNet` | `property, plant` |
| Equity investments | `EquityMethodInvestments`, `InvestmentsInAffiliatesSubsidiariesAssociatesAndJointVentures` | `investments in unconsolidated` |

Operating cost (history) = Total costs and expenses − D&A − Impairments −
Gains (where total costs and expenses is unmapped, Revenue − Operating income
− D&A − Impairments − Gains).

The builder completes the list; the pattern is what matters. Each concept
resolves to one as-reported row per year (it may be a different row in
different years, e.g. EPD's custom D&A tag in some years). The **Mapping**
sheet shows concept → tag → label → value per year, and marks `UNMAPPED` where
nothing matched. The historical model lines reference the as-reported rows by
formula, so an analyst can repoint a mapping in Excel without touching Python.

## 6. Workbook layout

All sheets: Calibri, blue font for hard-coded inputs, black for formulas,
yellow fill for user inputs, green fill for key outputs, $MM everywhere except
per-share. History columns first (oldest → newest), then forecast columns
`FY2026E`…`FY2030E` with a distinct fill. Column A margin, labels in B, first
year in D (C holds the row's unit/note). Freeze panes at the first year column.

1. **Cover** — company, CIK, FYE, years pulled, which filing supplied each
   year, check summary (counts of PASS/WARN/FAIL), sheet index with links,
   data-source note, date pulled.
2. **Checks** — the §5.5 table with values.
3. **Income Statement (As Reported)**, **Balance Sheet (As Reported)**,
   **Cash Flow (As Reported)** — every as-reported row, in order, with indent,
   section headers bold, subtotals bold with top border. Hard-coded SEC values
   (blue). Rows flagged "no longer reported" in italic gray.
4. **Mapping** — the §5.6 table. One row per concept per year? No: one row per
   concept, columns per year showing the value pulled, plus a text column with
   the tag and label used (most recent year), and a status column.
5. **Model** — the linked 3-statement model, history + forecast on one sheet
   in three blocks (IS / BS / CF) using the standard concept lines. History
   cells are formulas pointing at the As Reported sheets via the Mapping.
   Forecast cells are formulas off Drivers and Asset Inputs. See §7.
6. **Asset Inputs** — see §7.1.
7. **Drivers** — see §7.2.
8. **WACC** — see §8.
9. **DCF** — see §8.
10. **Sources** — per year per statement: accession number, filing date, form,
    R-file URL, scale text found. Plus the companyfacts API values used in the
    cross-check.

## 7. Forecast mechanics (Model sheet)

### 7.1 Asset Inputs
Rows: one per asset. Columns: Asset name, Segment (text), then Revenue by
forecast year, then Operating cost by forecast year. A final row per block is
"Other / unallocated". Below the blocks: Total revenue, Total operating cost,
Asset EBITDA (= revenue − cost), and a reconciliation row showing the last
reported year's consolidated revenue and operating cost next to the asset
totals, so an analyst sees the starting point they are picking up from.

Ships pre-filled with one asset, "Consolidated", whose forecast revenue grows
at the historical 3-year revenue CAGR (capped at ±15%) and whose cost holds
the last reported cost-to-revenue ratio. These are formulas off Drivers, so
replacing them with real asset rows is: type over the row, add rows, totals
pick them up (totals use whole-column SUM ranges with a fixed block size of 25
asset rows, enough for the demo).

### 7.2 Drivers
Yellow inputs, one column per forecast year, each defaulting to a formula off
the company's own history (3-year average unless noted), so nothing is a
magic number:
- D&A as % of opening net PP&E (fallback: % of revenue if PP&E unmapped)
- Capex as % of revenue
- Interest rate on opening total debt (= last year interest / avg debt)
- Effective tax rate (= tax / pretax income, clamped 0–35%; a "cash tax rate"
  that subtracts deferred taxes is shown as an information line only; partnerships will come
  out near 0)
- Distributions/dividends: per-unit amount per year (last declared × growth
  input, default growth = 3-year CAGR) × units outstanding
- Share/unit count: flat (input)
- Working capital: receivables days, inventory days, payables days off revenue
  / cost, defaults = last year. Kept in one small block so it is easy to
  ignore; a "Hold working capital flat" switch (yes/no, default no) zeroes the
  change.
- Debt: new issuance / repayment per year (input, default 0); minimum cash
  balance (default = last reported cash); revolver draws/repays automatically
  to hold minimum cash.
- Other income statement lines not covered (equity earnings, other
  income/expense, NI to noncontrolling): held flat at the last reported value.

### 7.3 Statement links (forecast columns)
- IS: Revenue and Operating cost from Asset Inputs → Asset EBITDA; plus
  equity earnings from unconsolidated affiliates (held flat) if the JV
  toggle is Yes → EBITDA (if No, equity earnings are added after EBITDA) → less
  D&A, impairments (0), gains (0) → EBIT → less interest (on opening debt +
  opening revolver, so no circularity) → plus other income/expense (flat) →
  pretax → tax → net income → less NI to noncontrolling → NI to
  common/partners. Operating cost definition is in §9.
- CF: NI + D&A − equity earnings + distributions received from affiliates
  (held flat at last reported) − Δworking capital = OCF; − capex =
  investing; − distributions to common/partners − distributions to NCI
  ± debt issuance/repayment ± revolver = financing; net change in cash;
  ending cash.
- BS: Cash = prior + net change. Net PP&E = prior + capex − D&A. Working
  capital items off days. Debt = prior ± issuance/repayment; revolver
  separate line. Equity = prior + NI to common − distributions. Everything
  else (other assets, other liabilities, noncontrolling interest) held flat.
  A **balance check row** = assets − (liabilities + equity), which must be
  0.00 in every forecast year by construction (cash is the balancing item via
  the CF). A conditional format turns it red if |value| > 0.5.

Historical columns of the Model sheet show the same standard lines, pulled by
formula from the As Reported sheets, so ratios are computed on the same basis
history and forecast.

## 8. WACC and DCF sheets

Follow `Desktop/Modeling/DCFWacc Templates.xlsx` structure, rebuilt with live
formulas and no `#REF!`:

**WACC**: cost of debt (input, default = Model's historical interest rate),
tax rate (from Model), after-tax Kd; Rf, beta, MRP inputs, Ke = Rf + β·MRP;
capital weights from market cap (price input × diluted units from Model) and
net debt (from Model, last year); target weight override; WACC. Beta
selector: Observed (input) or Industry (relevered from a peer table of up to 6
rows: ticker, observed beta, price, shares, net debt, tax rate → unlevered
beta → average → relevered to the subject). Peer rows are blank yellow inputs.
WACC sensitivity (beta × MRP) as a formula grid.

**DCF**: company header; valuation date = today; stub-year fraction; UFCF
build from Model forecast lines (EBITDA, EBIT, tax on EBIT, NOPAT, D&A,
ΔNWC, capex, UFCF); mid-year toggle; discount factors; PV of UFCF. Terminal
value both ways: perpetuity growth (input g, default 2.5%) and EBITDA exit
multiple (input, default = current EV / last-year EBITDA computed from the
price input). Implied multiple / implied growth cross-checks. Net debt bridge
(debt, preferred, NCI, less cash and equity investments, all from Model with
override cells). Equity value and value per share/unit under both methods,
and % over/under current price. Two sensitivity grids as formula grids:
value per share for WACC × g and WACC × exit multiple. A football-field table
(low / high per method plus 52-week range inputs) and a stacked bar chart
drawn by openpyxl.

## 9. Decisions made here (my calls, open to override)

- **New main app file `app.py`** (and `sec_model/` package: `edgar.py`,
  `parse.py`, `merge.py`, `checks.py`, `mapping.py`, `workbook.py`). The old
  `streamlit_app.py`, `sec_10k_app.py` and notebooks stay in the repo
  untouched until the new one is trusted, then get removed in a follow-up.
  README updated to point at `app.py`.
- **Market inputs are manual.** Share price, beta and 52-week range are yellow
  inputs with no live feed. SEC has no prices and a scraped quote is the first
  thing to break in a demo. Shares/units outstanding come from the 10-K cover
  (`dei:EntityCommonStockSharesOutstanding`).
- **10 years of history by default**, 5 forecast years. Both are sidebar
  inputs.
- **Forecast columns start the year after the latest 10-K.** 10-Q stub periods
  are not pulled; the DCF stub-year fraction handles the valuation-date
  offset, as in the template.
- **Interest on opening balances**, not average, to avoid circular references.
  Stated on the Drivers sheet.
- **Sensitivities are formula grids**, not Excel data tables. Each cell is a
  self-contained formula with the WACC/g/multiple substituted in; the DCF math
  is factored so that is possible (PV of UFCF and TV as a function of WACC in
  a helper block).
- **Disk cache** of SEC responses under a `.cache/` folder (gitignored), keyed
  by URL, with a 7-day expiry, so re-runs and tests are fast and polite.
- **Noncontrolling interest balances explicitly.** The CF financing section
  gets a "Distributions to noncontrolling interests" line (historical tag
  `PaymentsOfDividendsMinorityInterest`, label fallback `noncontrolling`).
  Forecast default = NI attributable to noncontrolling interests, so the BS
  NCI line stays flat and the model balances without a plug. Matters for ET
  (Sunoco, USAC), MPLX, WMB, KMI.
- **Market inputs: sidebar block, not blank cells.** The Streamlit sidebar
  gets a "Market inputs" expander: share/unit price (optional), beta (default
  1.0), risk-free rate (default 4.25%), market risk premium (default 5.5%).
  They are written into the WACC/DCF yellow input cells, so the workbook
  calculates on first open. If price is left blank the DCF still produces EV
  and value per unit; "% over/under current price" shows "enter price", the
  exit-multiple default falls back to 10.0x, and the price cell is flagged.
- **ENB and other non-USD filers: no conversion.** Currency is read from the
  R-file header (`CAD ($)`) and the companyfacts unit key. The workbook unit
  note says "CAD MM", every sheet title carries the currency, and the DCF
  shows a warning that price must be entered in the same currency. The
  companyfacts cross-check uses the detected currency unit.
- **Streamlit Cloud entry point.** `streamlit_app.py` becomes a two-line shim
  that imports and runs `app.py`, so the existing Cloud URL keeps working
  without a dashboard change. The old app logic is preserved in
  `sec_10k_app.py` (already identical) until removed in a follow-up.
- **Template deviations.** 5 forecast years, not 7. The DCF's diluted
  securities block (RSUs, options, convertibles, convertible preferred) is
  kept as zero-default yellow inputs added to basic units from the 10-K cover.
  The Current / Year-1 / Terminal multiples block is kept.
- **Operating cost definition (history and asset block).** "Operating cost"
  = every cost line between revenue and operating income **except** D&A,
  impairments, and gains/losses on asset sales. It therefore includes cost of
  product sold / commodity purchases, operations and maintenance, G&A, and
  taxes other than income. Impairments and gains are shown as their own
  non-cash lines below EBITDA and held at zero in the forecast, so a
  multi-billion impairment year (KMI, ET, TRGP) does not distort the default
  cost-to-revenue driver. Equity earnings from unconsolidated affiliates are
  a separate line next to EBITDA, and a **toggle on the Drivers sheet,
  "Include JV equity earnings in EBITDA" (Yes/No, default Yes)**, decides
  whether they sit inside the EBITDA subtotal. The IS shows "Asset EBITDA"
  (revenue − operating cost) and "EBITDA" (= Asset EBITDA + equity earnings
  × toggle); the DCF's EBITDA, exit multiple and EV/EBITDA multiples all
  reference the toggled EBITDA line. When the toggle is No, equity earnings
  are added below EBITDA before EBIT so net income is unchanged either way.
  Decided by the user on 2026-10-22.

## 12. Build order if time runs short

Must have for Wednesday, in order: (1) pipeline + as-reported sheets +
checks for OKE/EPD/ET/KMI/WMB/MPLX/TRGP/KNTK/ENB; (2) Model sheet with asset
block and balancing forecast; (3) WACC + DCF with both terminal methods and
the two value-per-unit sensitivity grids; (4) regression script over the
midstream list; (5) AAPL/XOM/JPM builds.
Cut first, in order: football-field chart (keep the table), peer-beta
relevering table (keep observed beta), bank-layout and non-December parser
fixtures (keep the midstream fixtures), working-capital "hold flat" switch.

## 10. Testing and acceptance

- Unit tests (pytest) for the R-file parser against saved HTML fixtures from
  at least: OKE (IS/BS/CF, `$ in Millions`), EPD (custom tags), ET (blank
  first column + footnote markers), one `$ in Thousands` filer, one non-December
  FYE filer, one bank-style filer (no "operating income"; e.g. JPM) to prove
  the parser does not assume an industrial layout.
- A regression run script `scripts/regress.py TICKERS...` that builds the
  workbook for each ticker and asserts: no exception; ≥ N years found; all
  three statements found for every year; every FAIL-level check passes; the
  Model sheet balance-check row is zero in every forecast year (evaluated in
  Python by recomputing the forecast, since openpyxl does not calculate); the
  workbook re-opens with openpyxl without error.
- Acceptance list for Wednesday: OKE, EPD, ET, KMI, WMB, MPLX, TRGP, KNTK, ENB
  all pass the regression script; AAPL, XOM, JPM also build without FAIL (to
  show it is not midstream-only); TRP returns the clean "40-F filer, not
  supported" message.
- Manual check: open OKE and EPD workbooks in Excel, no repair prompt, no
  `#REF!`, change one asset's revenue and confirm EBITDA, cash, debt and both
  DCF values move.

## 11. Left to the builder

- Exact keyword lists, label regexes and the full concept table in §5.6.
- Indent detection details in R-files (style vs nested div vs class).
- How to lay out the Model sheet rows so formulas stay readable (defined names
  for key rows are encouraged).
- Chart styling for the football field.
- Streamlit caching strategy (`st.cache_data` on the pipeline keyed by
  ticker + years).
