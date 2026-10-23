"""Download real SEC filings as test fixtures.

Usage: python scripts/fetch_fixtures.py OKE EPD ET --ua "Name email@x.com"

For each ticker: newest plain 10-K (10-K/A only if no 10-K), saved to
tests/fixtures/<TICKER>/ as FilingSummary.xml, IS.htm, BS.htm, CF.htm.
"""
import argparse
import html as htmllib
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from sec_model.edgar import list_annual_filings, locate_statements, resolve_ticker  # noqa: E402
from sec_model.http import SecClient  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"


def table_title(page: str) -> str:
    """Text of the first header cell of the first table (the '... - USD ($) $ in Millions' line)."""
    m = re.search(r"<th[^>]*>(.*?)</th>", page, re.S | re.I)
    if not m:
        m = re.search(r"<t[dh][^>]*>(.*?)</t[dh]>", page, re.S | re.I)
    if not m:
        return "(no table header found)"
    return " ".join(htmllib.unescape(re.sub(r"<[^>]+>", " ", m.group(1))).split())


def newest_10k(filings):
    plain = [f for f in filings if f.form == "10-K"]
    return (plain or filings)[-1]  # list_annual_filings sorts oldest -> newest


def fetch_one(client, ticker: str) -> None:
    cik, name = resolve_ticker(client, ticker)
    filings = list_annual_filings(client, cik)
    filing = newest_10k(filings)
    print(f"== {ticker}: {name} (CIK {cik}) {filing.form} FYE {filing.report_date} acc {filing.accession}")
    located = locate_statements(client, filing)
    out = FIXTURES / ticker.upper()
    out.mkdir(parents=True, exist_ok=True)
    (out / "FilingSummary.xml").write_bytes(client.get_bytes(filing.folder_url + "FilingSummary.xml"))
    for stype, rep in located.items():
        if rep is None:
            print(f"   {stype}: NOT FOUND")
            continue
        page = client.get_text(filing.folder_url + rep.html_file)
        (out / f"{stype}.htm").write_text(page, encoding="utf-8")
        print(f"   {stype}: {rep.html_file} [{rep.short_name}] -> {table_title(page)}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("tickers", nargs="+")
    ap.add_argument("--ua", required=True, help='SEC User-Agent, e.g. "Name email@x.com"')
    args = ap.parse_args()
    client = SecClient(args.ua)
    for t in args.tickers:
        try:
            fetch_one(client, t)
        except Exception as exc:  # keep going; report which ticker failed
            print(f"== {t}: FAILED {type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
