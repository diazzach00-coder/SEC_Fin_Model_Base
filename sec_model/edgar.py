"""Ticker -> CIK resolution and annual (10-K family) filing selection."""
from datetime import date

from sec_model.types import Filing, UserFacingError

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
TICKERS_EXCHANGE_URL = "https://www.sec.gov/files/company_tickers_exchange.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/{name}"

ANNUAL_FORMS = {"10-K", "10-K/A", "10-KT"}
FOREIGN_ANNUAL_FORMS = {"40-F", "40-F/A", "20-F", "20-F/A"}


def _norm(ticker: str) -> str:
    return ticker.strip().upper().replace(".", "-")


def resolve_ticker(client, ticker: str) -> tuple[int, str]:
    """Return (cik, company_name) for a ticker; '.' and '-' are treated as equal."""
    t = ticker.strip()
    target = _norm(t)

    for entry in client.get_json(TICKERS_URL).values():
        if _norm(entry["ticker"]) == target:
            return int(entry["cik_str"]), entry["title"]

    exchange = client.get_json(TICKERS_EXCHANGE_URL)
    fields = exchange["fields"]
    i_cik, i_name, i_ticker = fields.index("cik"), fields.index("name"), fields.index("ticker")
    for row in exchange["data"]:
        if _norm(row[i_ticker]) == target:
            return int(row[i_cik]), row[i_name]

    raise UserFacingError(f"Ticker '{t}' not found in SEC company list.")


def _iter_forms(cik: int, arrays: dict):
    """Yield (form, Filing-or-None) per entry; Filing is None when dates are missing."""
    for i, form in enumerate(arrays.get("form", [])):
        report, filed = arrays["reportDate"][i], arrays["filingDate"][i]
        if not report or not filed:
            yield form, None
            continue
        yield form, Filing(
            cik=cik,
            accession=arrays["accessionNumber"][i],
            form=form,
            filing_date=date.fromisoformat(filed),
            report_date=date.fromisoformat(report),
            primary_doc=arrays["primaryDocument"][i],
        )


def list_annual_filings(client, cik: int) -> list[Filing]:
    """All 10-K / 10-K/A / 10-KT filings (recent + paginated), oldest report first."""
    sub = client.get_json(SUBMISSIONS_URL.format(name=f"CIK{cik:010d}.json"))
    filings_block = sub["filings"]
    pages = [filings_block["recent"]]
    for f in filings_block.get("files", []):
        pages.append(client.get_json(SUBMISSIONS_URL.format(name=f["name"])))

    annual: list[Filing] = []
    foreign_forms: list[str] = []
    for page in pages:
        for form, filing in _iter_forms(cik, page):
            if form in FOREIGN_ANNUAL_FORMS:
                foreign_forms.append(form)
            elif form in ANNUAL_FORMS and filing is not None:
                annual.append(filing)

    if not annual and foreign_forms:
        raise UserFacingError(
            f"CIK {cik} files Form {foreign_forms[0]} (no US-GAAP XBRL statements); not supported."
        )
    annual.sort(key=lambda f: (f.report_date, f.filing_date))
    return annual


def select_filings_by_fy(filings: list[Filing], n_years: int) -> dict[str, list[Filing]]:
    """Group by fiscal year, newest filing first within a year; keep the n_years latest FYs."""
    by_fy: dict[str, list[Filing]] = {}
    for f in filings:
        by_fy.setdefault(f.fy_label, []).append(f)
    latest = sorted(by_fy)[-n_years:] if n_years > 0 else []
    return {
        fy: sorted(by_fy[fy], key=lambda f: f.filing_date, reverse=True) for fy in latest
    }
