from datetime import date

import pytest

from sec_model.edgar import list_annual_filings, resolve_ticker, select_filings_by_fy
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
