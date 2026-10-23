# SEC 3-Statement Model Builder

A Streamlit app that pulls as-reported 10-K financials from SEC EDGAR, ties them out,
and exports a linked 3-statement Excel model with WACC and DCF. No API key or data
subscription is needed; everything comes from SEC EDGAR.

## What it does

Enter a ticker and a User-Agent (name and email). The app then:

1. Resolves the ticker to a CIK and lists the company's 10-K / 10-K/A / 10-KT filings.
2. Reads the income statement, balance sheet and cash flow statement of each fiscal year
   from the filing's XBRL "R" pages (the as-reported presentation, in reported order).
3. Merges the filings into one history and maps the as-reported lines to the model's
   concepts. Each fiscal year's three statements come from one filing: the newest that
   covers that year; restatements that appear only in a later filing's comparatives are
   not used.
4. Runs data checks (below), builds a forecast, and writes the workbook.
5. Shows a progress log, the checks table, previews of each statement and the mapping,
   and a download button for `{TICKER}_3Stmt_{YYYY-MM-DD}.xlsx`.

## Workbook sheets

| Sheet | Contents |
|---|---|
| **Cover** | Company, CIK, fiscal year end, years pulled, filing used per year, check summary, sheet index, date pulled |
| **Checks** | Every data check with the numbers compared and the tolerance |
| **Model** | Linked income statement, balance sheet and cash flow, history plus forecast on one sheet. History links to the As-Reported sheets through Mapping; the forecast is formulas off Drivers and Asset Inputs |
| **Asset Inputs** | Asset-level revenue and operating cost by forecast year; row 1 is a placeholder that grows the consolidated figures, to be overwritten with real assets (the hand-off point for asset models) |
| **Drivers** | Forecast drivers (margins, working capital, capex, distributions, financing) with defaults from history, plus the JV-earnings toggle |
| **WACC** | CAPM cost of equity, cost of debt, capital structure, WACC |
| **DCF** | Unlevered cash flows, terminal value, EV to equity bridge, value per share/unit, and sensitivity grids built from ordinary formulas (not Excel Data Tables) |
| **Mapping** | One row per model concept: status, statement, value per year, and an editable row key so you can repoint a concept to another as-reported line |
| **IS / BS / CF (As Reported)** | Every as-reported line, in order, with indentation and the SEC's hard-coded values |
| **Sources** | Accession number, form, filing date and R-file for each year and statement |

## Data checks

Run on every build; a FAIL does not stop the build, it is flagged on the Cover sheet.

- **Balance sheet balances**: total assets equal total liabilities and equity.
- **Net income ties**: income statement net income equals the first line of operating cash flow.
- **Cash roll**: beginning cash plus net change equals ending cash.
- **CF cash vs BS cash**: warning only (restricted cash is often in the cash flow figure).
- **Subtotals**: children of a "Total ..." row sum to it (warning).
- **API cross-check**: revenue, net income, total assets and operating cash flow against the
  SEC companyfacts API (warning, FAIL above 5% difference).
- **Coverage**: each year has all three statements and the forecast anchors are mapped.

Tolerances are the larger of $1MM and 0.5% of the larger figure, to absorb rounding.
Checks that cannot be evaluated are shown as SKIP (hidden in the app by default).

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open `http://localhost:8501`. `streamlit_app.py` is a two-line shim that runs the same
app, so `streamlit run streamlit_app.py` (and the existing Streamlit Cloud entry point) works too.

To run the tests:

```bash
pip install -r requirements-dev.txt
python -m pytest tests -W error
```

Tests marked `network` use the real SEC data (cached under `.cache/`) and are skipped when
that folder is absent.

## Inputs

| Field | Description |
|---|---|
| **Ticker** | US-listed ticker, e.g. `OKE`, `EPD`, `KMI` |
| **SEC User-Agent (name and email)** | e.g. `Jane Doe jane@example.com`. Required by the SEC's [fair-access policy](https://www.sec.gov/os/accessing-edgar-data); must contain an `@` |
| **Years of history** | 2 to 15 fiscal years (default 10); limited to the years the company has full XBRL statements for |
| **Forecast years** | 3 to 7 (default 5) |
| **Market inputs** | Share / unit price (leave 0 to enter it later in the workbook), beta, risk-free rate, market risk premium. They are written into the WACC and DCF input cells |

Downloaded SEC pages are cached on disk in `.cache/`, so re-running a ticker is fast.

## Opening the workbook in Excel

The file is written with formulas and no cached values. Excel computes every formula when
the file is opened, so values appear immediately in Excel. Viewers that do not recalculate
(some previewers, Python readers such as `openpyxl` with `data_only=True`) will show empty cells.
Edit the yellow input cells (Asset Inputs, Drivers, WACC, DCF) and everything downstream updates.

## Limitations

- **Annual only.** Quarterly (10-Q) history is not built.
- **40-F and 20-F filers** (for example TC Energy, Pembina) have no US-GAAP XBRL statements;
  the app says so and stops.
- **Non-USD filers (for example Enbridge, CAD) are not converted.** The currency is detected
  from the filing, labelled on every sheet (for example "CAD MM"), and the share price you
  enter must be in the same currency.
- Only companies with XBRL "R" pages for all three statements in their 10-K are supported;
  years without them are dropped (the progress log says which).
- Where the filing's tagging is unusual, a concept can come out UNMAPPED. The Mapping sheet and
  the coverage warnings show where, and the row key can be repointed in Excel.
- Market inputs (price, beta, risk-free rate, MRP) are manual; no live market data is used.
- Other files in the repo (`sec_10k_app.py`, the notebooks) are earlier tools and are unchanged.

## Data source

SEC EDGAR, no API key:

- `https://www.sec.gov/files/company_tickers.json`: ticker to CIK
- `https://data.sec.gov/submissions/CIK{CIK}.json`: filing history
- `https://www.sec.gov/Archives/edgar/data/...`: `FilingSummary.xml` and the R pages of each filing
- `https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json`: cross-check values

## Built with

[Streamlit](https://streamlit.io), [pandas](https://pandas.pydata.org),
[openpyxl](https://openpyxl.readthedocs.io), [Beautiful Soup](https://www.crummy.com/software/BeautifulSoup/).

## Disclaimer

For educational and research purposes only. Not financial advice. Verify data against the
official SEC filings before making investment decisions.
