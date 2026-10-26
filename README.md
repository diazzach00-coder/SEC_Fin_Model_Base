# SEC 3-Statement Model Builder

Type a US stock ticker and get a linked, formula-driven Excel model built from the company's
own SEC filings: income statement, balance sheet and cash flow statement, a forecast, a cost
of capital (WACC) and a discounted cash flow (DCF) valuation. The app reads the as-reported
financial statements from each 10-K on SEC EDGAR, lines them up across years, checks that they
tie out, and writes the workbook. It works for any US-GAAP filer, whether a corporation
(shares, dividends) or a partnership (units, distributions); the words in the workbook follow
the entity type. No API key or paid data is needed.

The forecast is a mechanical starting point built from the company's history. It is meant to be
edited by an analyst, not trusted as it comes.

## Quick start

1. Clone the repository and open a terminal in it. You need Python 3.10 or newer.
2. Install the dependencies: `pip install -r requirements.txt`
3. Start the app: `streamlit run app.py`, then open `http://localhost:8501`.
4. In the sidebar type a **ticker** (for example `AAPL`) and your **name and email**
   (for example `Jane Doe jane@example.com`). The SEC requires requests to identify the person
   making them, so use your real details.
5. Click **Build model**, wait for the progress log to finish (a first run takes under a minute),
   then click **Download Excel model**. The file is named `{TICKER}_3Stmt_{YYYY-MM-DD}.xlsx`.

Other sidebar settings: years of history (2 to 15, default 10), forecast years (3 to 7,
default 5), a **Market inputs** box (share or unit price, beta, risk-free rate, market risk
premium; leave the price at 0 to enter it later in the workbook), and **CIK override
(advanced)** for the case described under Limitations. Downloaded SEC pages are cached in
`.cache/`, so running the same ticker again is fast.

## What you get

| Sheet | Contents |
|---|---|
| **Cover** | Company, CIK, fiscal year end, currency, years pulled, the filing used for each year, a count of check results, and links to every sheet |
| **Checks** | Every data check with the two numbers compared, the tolerance and a note |
| **Model** | Income statement, balance sheet and cash flow on one sheet, history then forecast, all linked, with ratios underneath and a balance check row |
| **Revenue Streams** | 25 rows each for revenue and operating cost by stream (segment, product, region, asset), by forecast year; row 1 is a placeholder that grows the company total |
| **Drivers** | Forecast assumptions by year (growth, margins, capex, tax, interest, working-capital days, payout, minimum cash) plus the model switches |
| **WACC** | CAPM cost of equity, cost of debt, capital weights, beta selection and a sensitivity grid; the share or unit price lives here |
| **DCF** | Unlevered free cash flow, terminal value, bridge to equity, value per share or unit, and sensitivity grids made of ordinary formulas (no Excel Data Tables) |
| **Mapping** | One row per model concept, with the as-reported line it was taken from in each year, and an editable key to repoint it |
| **Income Statement, Balance Sheet, Cash Flow (As Reported)** | Every line of the filing in its original order and indentation, with the SEC's numbers |
| **Sources** | Form, accession number, filing date and link to the filing page for each year and statement |

Colour legend, used on every sheet:

- **Blue text**: a number taken from the SEC data (reached through the Mapping sheet) or typed in.
- **Black text**: a formula.
- **Yellow fill**: an input you can overwrite.
- **Green fill**: a key result, such as value per share or unit.
- **Light-blue fill**: forecast columns.
- **Red**: a failed check or a cash shortfall.

The file has formulas and no stored results: Excel calculates everything on opening, but
viewers that do not calculate (some previewers, `openpyxl` with `data_only=True`) show empty cells.

## How to use the model

**Revenue Streams.** Overwrite row 1 ("Total company (placeholder)") or add your own streams
in the rows below it, in both the revenue block and the operating cost block, for each forecast
year. Blank rows count as zero. The Model sheet uses the two totals. Row 1 grows last year's
revenue by the Drivers revenue growth and sets cost as a percent of revenue, so with no edits
you get a consolidated forecast. A reconciliation block under the tables compares your
streams with the last reported year.

**Drivers.** Year 1 of each driver is a formula over the company's history (for example a
three-year average or last year's value); later years copy the year before. Overwrite any cell.
Revenue growth and cost percent drive only the placeholder row. Debt issuance and other
operating items start at zero. Minimum cash defaults to last year's cash, and the revolver
interest rate defaults to the term-debt rate.

**The switches** (at the bottom of the Drivers sheet; the first two are a 1/0 dropdown):

- **JV toggle** (`JV_TOGGLE`): 1 counts equity-method earnings inside EBITDA, which suits
  partnerships and joint-venture-heavy companies; 0 (default) keeps them below EBITDA.
- **Hold working capital flat** (`WC_HOLD`): 1 keeps receivables, inventory and payables at last
  year's balances so working capital does not move cash; 0 (default) uses the day-count drivers.
- **Revolver commitment** (`REV_CAP`): the most the revolver can borrow. The revolver draws
  to keep cash at the minimum cash driver and repays when cash is above it. 0 (default) means
  unlimited. If a commitment is too small, cash falls below the minimum and the shortfall
  row on the Model sheet turns red.
- **Cap on default growth rates** (`GROWTH_CAP`, default 15%) limits the default revenue and
  payout growth drivers.

**Price and beta.** On the WACC sheet, enter the share or unit price in the yellow box in the
right-hand block, and the risk-free rate, market risk premium and beta on the left (the beta
used is chosen by "Select beta for model"). The DCF sheet reads them from there. With no price
the WACC uses a 55% equity weight, which you can change, and the value-versus-price line says
"enter price".

**Reading the Checks sheet.** Filter the Status column. PASS is fine, WARN is worth a look
and FAIL means two figures that should agree do not. A failed check does not stop the build:
it is counted on the Cover sheet and shown in red. SKIP means a check could not be evaluated
(an input was missing); SKIPs are not problems by themselves.

**Repointing a Mapping key.** Each concept (for example Revenue or Receivables) is matched to
a line of the filing by its XBRL tag, or by its label when the tag is unusual. If the match is
wrong or missing, go to the Mapping sheet, find the concept, and in the yellow "Key" cell of the
year paste the row key of the right line from column A of the matching As-Reported sheet. The
value and everything downstream update. A status of UNMAPPED means no line was found.

## Checks

Run on every build for each fiscal year. Differences are measured against a tolerance of the
larger of 1 (in millions) or 0.5% of the larger figure, to allow for rounding.

| Check | What it tests | Level when it fails |
|---|---|---|
| BS balances | Total assets equal total liabilities and equity | FAIL |
| NI ties IS=CF | Net income on the income statement equals the first line of the cash flow statement | FAIL |
| Cash roll-forward | Opening cash plus the net change equals ending cash; opening is the prior year's ending cash | FAIL (WARN if only the cash definition changed between filings) |
| CF cash vs BS cash | Cash flow ending cash equals balance sheet cash | WARN (restricted cash is often included) |
| Subtotal | The indented lines above a "Total ..." row add up to it | WARN |
| API cross-check | Revenue, net income, total assets and operating cash flow against the SEC's companyfacts data | WARN, FAIL if more than 5% apart |
| Coverage | All three statements are present each year and the key lines (revenue, operating income, D&A, cash, total debt) are mapped | WARN |

## Limitations

- **Annual 10-K filings only.** Quarterly filings are not used.
- **Foreign filers (Forms 40-F and 20-F) are not supported.** They have no US-GAAP XBRL
  statements; the app says so and stops.
- **Non-US-dollar reporters are shown unconverted.** The currency is read from the filing and
  labelled on every sheet (for example "CAD MM"); any price you enter must use that currency.
- **Banks are not modelled.** Their statements have no revenue or receivables in the usual
  sense, so the forecast and DCF are not meaningful.
- **Predecessor history.** After a merger or reorganisation, older years may belong to the
  predecessor company. Each year is consistent on its own, but growth across the break is not
  meaningful.
- **A new holding company can have no filings yet.** If a ticker now points to a successor
  that has filed no 10-K (ExxonMobil was one), find the old company's CIK number on EDGAR and
  enter it in "CIK override (advanced)".
- **Each year's three statements come from one filing**: the newest that covers all three for
  that year. A restatement that appears only in a later filing's comparative columns is not used.
- **Subtotal checks are mostly SKIP**, because many statements do not indent the lines under a
  total. Treat the other checks as the real tests.
- **Unusual tagging can leave a concept UNMAPPED.** The Mapping sheet and the coverage
  warnings show where; repoint it as described above.
- Market inputs are typed in by hand; no live market data is used.

## For developers

```bash
pip install -r requirements-dev.txt
python -m pytest tests -W error
```

Tests marked `network` read real SEC pages from the `.cache/` folder and are skipped when it is
absent. The code is in `sec_model/` (reading and merging filings, `mapping.py` and
`concepts.py` for the concept matching, `model_spec.py` for the model written once as
formulas, `checks.py`, and `workbook/` for the Excel sheets); `app.py` is the Streamlit
interface and `streamlit_app.py` a one-line entry point for hosting.

The regression script builds live workbooks for a list of tickers and asserts sanity limits
(balance, interest and tax rates, working capital, unit counts); add `--excel` to recalculate
each file in Excel on Windows and check for error cells:

```bash
python scripts/regress.py --ua "Jane Doe jane@example.com" AAPL CVX OKE
python scripts/regress.py --ua "Jane Doe jane@example.com" --cik XOM=34088 XOM
```

Design and plans are in `docs/superpowers/`: the design spec in `specs/`, and the build plans
and the acceptance-run results (per-ticker results and known limitations) in `plans/`.

## Data source and fair use

All data comes from SEC EDGAR, free and without a key: the ticker list
(`sec.gov/files/company_tickers.json`), each company's filing list (`data.sec.gov/submissions`),
the financial-report pages of each 10-K (`sec.gov/Archives/edgar/data`), and the companyfacts
API (`data.sec.gov/api/xbrl/companyfacts`) for cross-checks only. The app follows the SEC's
[fair-access policy](https://www.sec.gov/os/accessing-edgar-data): it sends your name and
email as the User-Agent, limits its request rate and caches pages on disk. Please keep it that
way if you change the code.

Built with Streamlit, pandas, openpyxl and Beautiful Soup.

For education and research only; not financial advice. Check the data against the official
filings before relying on it.
