from datetime import date
from pathlib import Path

import pytest
import requests

from sec_model.edgar import (
    ReportEntry,
    choose_source_filing,
    list_annual_filings,
    locate_statements,
    parse_filing_summary,
    pick_statement_reports,
    resolve_ticker,
    select_filings_by_fy,
)
from sec_model.types import Filing, UserFacingError

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
EXCHANGE_URL = "https://www.sec.gov/files/company_tickers_exchange.json"


def sub_url(name: str) -> str:
    return f"https://data.sec.gov/submissions/{name}"


class FakeClient:
    def __init__(self, payloads: dict):
        self.payloads = payloads

    def get_json(self, url: str) -> dict:
        return self.payloads[url]


def arrays(rows):
    """rows: (accession, filingDate, reportDate, form, primaryDocument)"""
    keys = ["accessionNumber", "filingDate", "reportDate", "form", "primaryDocument"]
    return {k: [r[i] for r in rows] for i, k in enumerate(keys)}


def make_filing(form, filed, report, acc="0000000001-00-000001"):
    return Filing(1, acc, form, filed, report, "doc.htm")


def test_resolve_ticker_case_and_class_suffix():
    fake = FakeClient({TICKERS_URL: {
        "0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."},
        "1": {"cik_str": 1067983, "ticker": "BRK-B", "title": "BERKSHIRE HATHAWAY INC"},
    }})
    assert resolve_ticker(fake, "brk.b") == (1067983, "BERKSHIRE HATHAWAY INC")
    assert resolve_ticker(fake, " aapl ") == (320193, "Apple Inc.")


def test_resolve_ticker_falls_back_to_exchange_file():
    fake = FakeClient({
        TICKERS_URL: {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}},
        EXCHANGE_URL: {
            "fields": ["cik", "name", "ticker", "exchange"],
            "data": [[1067983, "BERKSHIRE HATHAWAY INC", "BRK-B", "NYSE"]],
        },
    })
    assert resolve_ticker(fake, "BRK.B") == (1067983, "BERKSHIRE HATHAWAY INC")


def test_resolve_ticker_not_found():
    fake = FakeClient({
        TICKERS_URL: {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}},
        EXCHANGE_URL: {"fields": ["cik", "name", "ticker", "exchange"], "data": []},
    })
    with pytest.raises(UserFacingError, match="Ticker 'zzzz' not found in SEC company list."):
        resolve_ticker(fake, "zzzz")


def test_list_annual_filings_reads_paginated_files():
    fake = FakeClient({
        sub_url("CIK0001039684.json"): {"filings": {
            "recent": arrays([
                ("0001039684-26-000012", "2026-02-20", "2025-12-31", "10-K", "a.htm"),
                ("0001039684-26-000005", "2026-01-10", "2025-12-31", "8-K", "b.htm"),
                ("0001039684-25-000020", "2025-05-01", "2025-03-31", "10-Q", "c.htm"),
                ("0001039684-25-000021", "2025-05-01", "", "10-K/A", "d.htm"),
            ]),
            "files": [{"name": "CIK0001039684-submissions-001.json",
                       "filingCount": 1, "filingFrom": "2015-01-01", "filingTo": "2016-12-31"}],
        }},
        sub_url("CIK0001039684-submissions-001.json"): arrays([
            ("0001039684-16-000003", "2016-02-25", "2015-12-31", "10-K", "e.htm"),
        ]),
    })
    fs = list_annual_filings(fake, 1039684)
    assert [f.fy_label for f in fs] == ["FY2015", "FY2025"]
    assert fs[0].accession == "0001039684-16-000003"
    assert fs[0].cik == 1039684
    assert fs[1].filing_date == date(2026, 2, 20)
    assert fs[1].primary_doc == "a.htm"


def test_list_annual_filings_includes_amendments_and_transition_sorted():
    fake = FakeClient({sub_url("CIK0000000042.json"): {"filings": {
        "recent": arrays([
            ("a3", "2024-09-01", "2023-12-31", "10-K/A", "x.htm"),
            ("a1", "2024-03-01", "2023-12-31", "10-K", "y.htm"),
            ("a2", "2023-03-01", "2022-06-30", "10-KT", "z.htm"),
        ]),
        "files": [],
    }}})
    fs = list_annual_filings(fake, 42)
    assert [(f.form, f.accession) for f in fs] == [
        ("10-KT", "a2"), ("10-K", "a1"), ("10-K/A", "a3"),
    ]


def test_40f_filer_raises():
    fake = FakeClient({sub_url("CIK0001232384.json"): {"filings": {
        "recent": arrays([
            ("t1", "2025-02-20", "2024-12-31", "40-F", "a.htm"),
            ("t2", "2025-05-01", "2025-03-31", "6-K", "b.htm"),
        ]),
        "files": [],
    }}})
    with pytest.raises(UserFacingError, match="40-F"):
        list_annual_filings(fake, 1232384)


def test_20f_amendment_named_in_message():
    fake = FakeClient({sub_url("CIK0000000007.json"): {"filings": {
        "recent": arrays([("t1", "2025-02-20", "2024-12-31", "20-F/A", "a.htm")]),
        "files": [],
    }}})
    with pytest.raises(UserFacingError, match="20-F/A"):
        list_annual_filings(fake, 7)


def test_10k_present_alongside_40f_does_not_raise():
    fake = FakeClient({sub_url("CIK0000000008.json"): {"filings": {
        "recent": arrays([
            ("t1", "2020-02-20", "2019-12-31", "40-F", "a.htm"),
            ("t2", "2025-02-20", "2024-12-31", "10-K", "b.htm"),
        ]),
        "files": [],
    }}})
    assert [f.form for f in list_annual_filings(fake, 8)] == ["10-K"]


def test_no_annual_filings_returns_empty():
    fake = FakeClient({sub_url("CIK0000000009.json"): {"filings": {
        "recent": arrays([("t1", "2025-02-20", "2024-12-31", "8-K", "a.htm")]),
        "files": [],
    }}})
    assert list_annual_filings(fake, 9) == []


def test_select_filings_by_fy_groups_amendments_newest_first():
    filings = [
        make_filing("10-K", date(2022, 3, 1), date(2021, 12, 31), "a0"),
        make_filing("10-K", date(2023, 3, 1), date(2022, 12, 31), "a1"),
        make_filing("10-K", date(2024, 3, 1), date(2023, 12, 31), "a2"),
        make_filing("10-K", date(2025, 3, 1), date(2024, 12, 31), "a3"),
        make_filing("10-K/A", date(2025, 8, 1), date(2024, 12, 31), "a4"),
    ]
    sel = select_filings_by_fy(filings, n_years=2)
    assert list(sel) == ["FY2023", "FY2024"]
    assert [f.form for f in sel["FY2024"]] == ["10-K/A", "10-K"]
    assert [f.accession for f in sel["FY2023"]] == ["a2"]


def test_select_filings_by_fy_fewer_years_than_requested():
    filings = [make_filing("10-K", date(2025, 3, 1), date(2024, 12, 31))]
    assert list(select_filings_by_fy(filings, n_years=5)) == ["FY2024"]


# ---- Task 3: locating the three statements -------------------------------
FIXTURES = Path(__file__).parent / "fixtures"


def FIX(rel: str) -> str:
    return (FIXTURES / rel).read_text(encoding="utf-8")


def summary_xml(reports):
    """reports: (short, category, html_file or None)"""
    body = ""
    for short, cat, html in reports:
        tag = f"<HtmlFileName>{html}</HtmlFileName>" if html else "<XmlFileName>R1.xml</XmlFileName>"
        body += (f"<Report><ShortName>{short}</ShortName><LongName>0001 - {short}</LongName>"
                 f"<MenuCategory>{cat}</MenuCategory>{tag}</Report>")
    return f'<?xml version="1.0"?><FilingSummary><MyReports>{body}</MyReports></FilingSummary>'


class TextClient:
    """get_text over a dict; a missing url behaves like an SEC 404."""

    def __init__(self, pages: dict):
        self.pages = pages

    def get_text(self, url: str) -> str:
        try:
            return self.pages[url]
        except KeyError:
            raise requests.HTTPError(f"404 for {url}")


def test_parse_filing_summary_skips_entries_without_html():
    xml = summary_xml([("Cover", "Cover", "R1.htm"), ("Old", "Statements", None)])
    assert parse_filing_summary(xml) == [ReportEntry("Cover", "0001 - Cover", "Cover", "R1.htm")]


def test_parse_filing_summary_tolerates_bom_and_encoding_declaration():
    xml = "﻿" + summary_xml([("Cover", "Cover", "R1.htm")]).replace(
        '<?xml version="1.0"?>', '<?xml version="1.0" encoding="utf-8"?>')
    assert [r.html_file for r in parse_filing_summary(xml)] == ["R1.htm"]


def test_pick_statement_reports_oke():
    reps = parse_filing_summary(FIX("OKE/FilingSummary.xml"))
    best = {k: v[0].html_file for k, v in pick_statement_reports(reps).items()}
    assert best == {"IS": "R3.htm", "BS": "R7.htm", "CF": "R9.htm"}


def test_pick_statement_reports_epd_bs_before_is():
    reps = parse_filing_summary(FIX("EPD/FilingSummary.xml"))
    best = {k: v[0].html_file for k, v in pick_statement_reports(reps).items()}
    assert best == {"IS": "R5.htm", "BS": "R3.htm", "CF": "R7.htm"}


def test_pick_excludes_parenthetical_and_comprehensive():
    reps = [ReportEntry("CONSOLIDATED STATEMENTS OF INCOME (Parenthetical)", "", "Statements", "R4.htm"),
            ReportEntry("CONSOLIDATED STATEMENTS OF COMPREHENSIVE INCOME", "", "Statements", "R5.htm"),
            ReportEntry("CONSOLIDATED STATEMENTS OF INCOME", "", "Statements", "R3.htm")]
    assert [r.html_file for r in pick_statement_reports(reps)["IS"]] == ["R3.htm"]


def test_pick_ignores_non_statement_categories_and_drops_zero_scores():
    reps = [ReportEntry("Consolidated Balance Sheets", "", "Details", "R9.htm"),
            ReportEntry("Document and Entity Information", "", "Cover", "R1.htm"),
            ReportEntry("Consolidated Statements of Cash Flows", "", "Statements", "R8.htm"),
            ReportEntry("Consolidated Statements of Equity", "", "Statements", "R6.htm")]
    picked = pick_statement_reports(reps)
    assert picked == {"IS": [], "BS": [], "CF": [reps[2]]}


def test_pick_precise_name_outranks_broad_name():
    reps = [ReportEntry("Income Taxes Summary", "", "Statements", "R2.htm"),
            ReportEntry("Statements of Earnings", "", "Statements", "R3.htm")]
    assert [r.html_file for r in pick_statement_reports(reps)["IS"]] == ["R3.htm", "R2.htm"]


IS_HTML = '<tr onclick="top.Show.showAR( this, \'defref_us-gaap_Revenues\', window )">'
BS_HTML = '<tr onclick="top.Show.showAR( this, \'defref_us-gaap_Assets\', window )">'
CF_HTML = '<tr onclick="x \'defref_us-gaap_NetCashProvidedByUsedInOperatingActivities\'">'


def make_filing_at(acc, form="10-K"):
    return Filing(7, acc, form, date(2026, 2, 20), date(2025, 12, 31), "doc.htm")


def folder(f):
    return f.folder_url


def test_locate_statements_picks_first_candidate_containing_anchor():
    f = make_filing_at("0000000007-26-000001")
    xml = summary_xml([
        ("Consolidated Statements of Income", "Statements", "R2.htm"),   # no anchor -> skipped
        ("Statements of Earnings", "Statements", "R3.htm"),
        ("Consolidated Balance Sheets", "Statements", "R4.htm"),
        ("Consolidated Statements of Cash Flows", "Statements", "R5.htm"),
    ])
    client = TextClient({
        folder(f) + "FilingSummary.xml": xml,
        folder(f) + "R2.htm": "<html>nothing here</html>",
        folder(f) + "R3.htm": IS_HTML,
        folder(f) + "R4.htm": BS_HTML,
        folder(f) + "R5.htm": CF_HTML,
    })
    located = locate_statements(client, f)
    assert {k: v.html_file for k, v in located.items()} == {"IS": "R3.htm", "BS": "R4.htm", "CF": "R5.htm"}


def test_locate_statements_404_filing_summary_gives_all_none():
    f = make_filing_at("0000000007-26-000002")
    assert locate_statements(TextClient({}), f) == {"IS": None, "BS": None, "CF": None}


def test_locate_statements_missing_r_file_is_skipped_not_fatal():
    f = make_filing_at("0000000007-26-000003")
    xml = summary_xml([("Consolidated Balance Sheets", "Statements", "R2.htm"),
                       ("Balance Sheets Alt", "Statements", "R3.htm")])
    client = TextClient({folder(f) + "FilingSummary.xml": xml, folder(f) + "R3.htm": BS_HTML})
    located = locate_statements(client, f)
    assert located["BS"].html_file == "R3.htm" and located["IS"] is None and located["CF"] is None


def test_choose_source_filing_skips_part3_amendment():
    amend = make_filing_at("0000000007-26-000010", "10-K/A")
    original = make_filing_at("0000000007-26-000005", "10-K")
    xml_full = summary_xml([
        ("Consolidated Statements of Income", "Statements", "R2.htm"),
        ("Consolidated Balance Sheets", "Statements", "R3.htm"),
        ("Consolidated Statements of Cash Flows", "Statements", "R4.htm"),
    ])
    client = TextClient({
        folder(amend) + "FilingSummary.xml": summary_xml([("Cover Page", "Cover", "R1.htm")]),
        folder(original) + "FilingSummary.xml": xml_full,
        folder(original) + "R2.htm": IS_HTML,
        folder(original) + "R3.htm": BS_HTML,
        folder(original) + "R4.htm": CF_HTML,
    })
    f, located = choose_source_filing(client, [amend, original])
    assert f is original and all(located.values())


def test_choose_source_filing_none_when_no_filing_has_all_three():
    amend = make_filing_at("0000000007-26-000010", "10-K/A")
    client = TextClient({folder(amend) + "FilingSummary.xml": summary_xml([("Cover", "Cover", "R1.htm")])})
    assert choose_source_filing(client, [amend]) is None
