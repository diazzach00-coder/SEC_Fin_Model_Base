"""Ticker -> CIK resolution, annual filing selection, and statement location."""
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date

import requests

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


# ---- Locating the three statements in a filing ----------------------------
@dataclass(frozen=True)
class ReportEntry:
    short_name: str
    long_name: str
    category: str
    html_file: str


_EXCLUDE = ("parenthetical", "comprehensive", "equity", "partners' capital",
            "details", "stockholders", "changes in")

_PRECISE = {
    "IS": ["statement of income", "statements of income", "statement of operations",
           "statements of operations", "statement of earnings", "statements of earnings",
           "statements of comprehensive income", "statements of consolidated operations"],
    "BS": ["balance sheet", "statements of condition", "financial position", "consolidated balance"],
    "CF": ["cash flow", "statements of cash", "cash flows", "statements of consolidated cash"],
}
_BROAD = {
    "IS": ["income", "operations", "earnings"],
    "BS": ["balance"],
    "CF": ["cash"],
}
_ANCHORS = {
    # SalesRevenueNet: the pre-ASC 606 revenue element (AAPL's FY2016-FY2017 10-Ks).
    "IS": ("us-gaap_Revenues", "us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax",
           "us-gaap_SalesRevenueNet", "us-gaap_CostsAndExpenses",
           "us-gaap_InterestAndDividendIncomeOperating"),
    "BS": ("us-gaap_Assets",),
    "CF": ("us-gaap_NetCashProvidedByUsedInOperatingActivities",),
}


def parse_filing_summary(xml: str) -> list[ReportEntry]:
    """Reports from FilingSummary.xml; entries without an HtmlFileName are skipped."""
    root = ET.fromstring(xml.lstrip("﻿").encode("utf-8"))
    out: list[ReportEntry] = []
    for rep in root.iter("Report"):
        html = (rep.findtext("HtmlFileName") or "").strip()
        if not html:
            continue
        out.append(ReportEntry(
            short_name=(rep.findtext("ShortName") or "").strip(),
            long_name=(rep.findtext("LongName") or "").strip(),
            category=(rep.findtext("MenuCategory") or "").strip(),
            html_file=html,
        ))
    return out


def _score(name: str, stype: str) -> int:
    precise = sum(kw in name for kw in _PRECISE[stype])
    broad = sum(kw in name for kw in _BROAD[stype])
    return precise * 10 + broad


# Some filers' XBRL puts a primary statement under "Uncategorized" (OKE's FY2021 10-K
# income statement); name scoring plus the anchor-tag check in locate_statements guard it.
_CATEGORIES = ("Statements", "Uncategorized")


def pick_statement_reports(reports: list[ReportEntry]) -> dict[str, list[ReportEntry]]:
    """Candidate reports per statement type, best first: higher name score, then
    "Statements" before "Uncategorized", then filing order."""
    eligible = []
    for r in reports:
        name = r.short_name.lower()
        if r.category in _CATEGORIES and not any(x in name for x in _EXCLUDE):
            eligible.append((r, name))
    picked: dict[str, list[ReportEntry]] = {}
    for stype in ("IS", "BS", "CF"):
        scored = [(_score(name, stype), r) for r, name in eligible]
        scored = [(s, r) for s, r in scored if s > 0]
        scored.sort(key=lambda sr: (-sr[0], _CATEGORIES.index(sr[1].category)))  # stable
        picked[stype] = [r for _, r in scored]
    return picked


def locate_statements(client, filing: Filing) -> dict[str, ReportEntry | None]:
    """The R-file for each statement: first candidate whose HTML carries the anchor tag."""
    none = {"IS": None, "BS": None, "CF": None}
    try:
        xml = client.get_text(filing.folder_url + "FilingSummary.xml")
    except requests.HTTPError:
        return none
    candidates = pick_statement_reports(parse_filing_summary(xml))
    located: dict[str, ReportEntry | None] = dict(none)
    for stype, cands in candidates.items():
        for rep in cands:
            try:
                html = client.get_text(filing.folder_url + rep.html_file)
            except requests.HTTPError:
                continue
            if any(a in html for a in _ANCHORS[stype]):
                located[stype] = rep
                break
    return located


def choose_source_filing(client, candidates: list[Filing]):
    """First filing (newest first) that yields all three statements, else None."""
    for filing in candidates:
        located = locate_statements(client, filing)
        if all(located.values()):
            return filing, located
    return None
