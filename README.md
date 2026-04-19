# SEC 10-K Financial Model Builder

A Streamlit web app that pulls SEC XBRL company facts directly from the SEC EDGAR API and builds a formatted, multi-sheet Excel financial model — no Bloomberg, no data subscriptions required.

---

## What It Does

Enter any US public company ticker and the app will:

1. Look up the company's CIK from the SEC ticker registry
2. Pull the 5 most recent 10-K filings
3. Download XBRL company facts (Income Statement, Balance Sheet, Cash Flow Statement)
4. Build a formatted Excel workbook with:
   - **Cover sheet** — company info, fiscal year summary, sheet index
   - **Income Statement** — revenue through EPS, margins, YoY growth
   - **Balance Sheet** — current/non-current assets, liabilities, equity, key ratios
   - **Cash Flow Statement** — operating/investing/financing flows, free cash flow
   - **WACC** — CAPM cost of equity, cost of debt, capital structure, sensitivity table
   - **Raw SEC Facts** — full XBRL tag audit trail as a formatted table
5. Offer a one-click download of the finished `.xlsx` file

All historical values are in **$MM** unless noted (EPS is per share; shares are in millions).

---

## Live App

> (https://financial-model-tools-dva8f8bk5rcbzxakungset.streamlit.app/)

---

## Files

| File | Purpose |
|---|---|
| `sec_streamlit_app.py` | Main Streamlit app |
| `requirements.txt` | Python package dependencies |

---

## Running Locally

```bash
pip install -r requirements.txt
streamlit run sec_streamlit_app.py
```

Then open `http://localhost:8501` in your browser.

---

## Inputs

| Field | Description |
|---|---|
| **Ticker** | US stock ticker, e.g. `AAPL`, `MSFT`, `TSLA` |
| **SEC User-Agent** | Your name and email, e.g. `Jane Doe jane@example.com`. Required by the SEC's fair-access policy for API requests. |

---

## Data Source

All data is fetched live from the **SEC EDGAR XBRL API** — no API key required:

- `https://www.sec.gov/files/company_tickers.json` — ticker-to-CIK lookup
- `https://data.sec.gov/submissions/CIK{CIK}.json` — filing history
- `https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json` — financial facts

> Per the SEC's [fair access policy](https://www.sec.gov/os/accessing-edgar-data), requests must include a descriptive `User-Agent` header with your name and contact email.

---

## Notes & Limitations

- Only companies that file **10-K** or **10-K/A** forms with XBRL data are supported (virtually all US public companies since ~2009).
- Some smaller or foreign-private issuers may have missing line items if their XBRL tagging is incomplete; the app reports any missing items clearly.
- Forecast columns (`FY[year]E`) are left blank for the user to fill in manually in Excel.
- The WACC sheet contains default input values — update risk-free rate, beta, ERP, market cap, and net debt for each company.

---

## Built With

- [Streamlit](https://streamlit.io)
- [pandas](https://pandas.pydata.org)
- [openpyxl](https://openpyxl.readthedocs.io)
- [SEC EDGAR API](https://www.sec.gov/developer)

---

## Disclaimer

This tool is for educational and research purposes only. It is not financial advice. Always verify data against official SEC filings before making investment decisions.
