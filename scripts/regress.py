"""Regression run: build the workbook for each ticker against live SEC data and assert
the acceptance rules (spec section 10).

    python scripts/regress.py --ua "Name email@example.com" OKE EPD ET ... TRP

Per ticker: writes ``<out>/<TICKER>.xlsx`` and prints one line (years found, statements
per year, check counts, coverage warnings, max |BAL_CHECK| over the forecast years,
elapsed). Assertion failures are collected; the exit code is 1 if any ticker failed.

Assertions: no exception, except a ``UserFacingError`` mentioning ``40-F`` for the 40-F
filers in ``EXPECT_40F``; at least ``--min-years`` fiscal years; IS, BS and CF for every
kept year; zero FAIL checks; max |BAL_CHECK| < 1e-6; the file reopens with openpyxl;
no ``#REF!`` in any cell.

Sanity assertions: default interest rate on debt within 0-15%; default cash tax rate
within 0-40%; last-year diluted units (UNITS) within 0.5x-2x of the dei shares
outstanding when both exist; no expense concept whose latest mapped value has the opposite
sign to the companyfacts value for its tag and year. NOTE lines (not asserted): interest
expense / pretax income unmapped in the latest year.

``--excel`` recalculates each workbook in Excel (COM), asserts no error cells and prints
the DCF value per unit (perpetuity / exit) for eyeballing. ``--cik TICKER=CIK`` builds a
ticker with a CIK override. ``KNOWN_FAILURES`` lists documented failures: reported as
KNOWN and not counted in the exit code while the failure text still matches.
"""
import argparse
import logging
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from sec_model.checks import summarize  # noqa: E402
from sec_model.companyfacts import fy_values, load_companyfacts  # noqa: E402
from sec_model.http import SecClient  # noqa: E402
from sec_model.concepts import CONCEPTS  # noqa: E402
from sec_model.pipeline import STYPES, _basic_units, build  # noqa: E402
from sec_model.types import UserFacingError  # noqa: E402

EXPECT_40F = {"TRP"}
BAL_TOL = 1e-6
# Model convention: these are positive when they reduce income (NIToNCI: positive = NCI's
# share of income), which is also the XBRL sign of their us-gaap tags.
SIGN_CONCEPTS = ("TotalCostsAndExpenses", "InterestExpense", "IncomeTax", "NIToNCI", "DA")
INT_RATE_RANGE, TAX_RATE_RANGE, UNITS_RATIO = (0.0, 0.15), (0.0, 0.40), (0.5, 2.0)
# ticker -> text the failure must contain. XOM: since 2026-07-01 the ticker maps to
# ExxonMobil Holdings Corp (CIK 2115436), a successor issuer with no 10-K yet; the history
# is under Exxon Mobil Corp, CIK 34088 (``--cik XOM=34088`` builds it).
KNOWN_FAILURES = {"XOM": "No 10-K filings found for XOM"}
# Banks: interest expense is mostly on deposits, so interest / total debt is not a cost of
# debt and the INT_RATE range is reported as a NOTE instead of asserted. (The model is not
# built for banks; JPM is in the list to show the build does not assume an industrial IS.)
BANKS = {"JPM"}
XL_ERRORS = {-2146826288: "#NULL!", -2146826281: "#DIV/0!", -2146826273: "#VALUE!",
             -2146826265: "#REF!", -2146826259: "#NAME?", -2146826252: "#NUM!",
             -2146826246: "#N/A"}


@dataclass
class Row:
    ticker: str
    years: str = ""
    n_years: int = 0
    stmts: str = ""
    counts: dict = field(default_factory=dict)
    coverage: list = field(default_factory=list)
    bal: float | None = None
    secs: float = 0.0
    kb: int = 0
    currency: str = ""
    note: str = ""
    failures: list = field(default_factory=list)
    sanity: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    vpu: str = ""
    known: bool = False

    @property
    def status(self) -> str:
        return "KNOWN" if self.known else "FAIL" if self.failures else "ok"


def scan_ref_errors(path: Path) -> tuple[list[str], str | None]:
    """Cells holding '#REF!' (value or formula text), and an openpyxl load error if any."""
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as e:  # any load failure is an assertion failure, not a crash
        return [], f"{type(e).__name__}: {e}"
    hits = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "#REF!" in c.value:
                    hits.append(f"{ws.title}!{c.coordinate}")
    return hits, None


def _first(v):
    return v[0] if isinstance(v, list) and v else v


def sanity_checks(res, ua: str) -> tuple[list[str], list[str]]:
    """(asserted sanity failures, informational notes) on model semantics."""
    bad, notes = [], []
    for key, (lo, hi) in (("INT_RATE", INT_RATE_RANGE), ("TAX_RATE", TAX_RATE_RANGE)):
        v = _first(res.driver_defaults.get(key))
        if isinstance(v, (int, float)) and not lo <= v <= hi:
            msg = f"default {key} {v:.2%} outside {lo:.0%}-{hi:.0%}"
            if key == "INT_RATE" and res.ticker in BANKS:
                notes.append(msg + " (bank: deposit interest; not asserted)")
            else:
                bad.append(msg)
    latest = res.fy_labels[-1]
    for concept in ("InterestExpense", "PretaxIncome"):
        cv = res.mapping.get(concept, {}).get(latest)
        if cv is None or cv.status == "UNMAPPED":
            notes.append(f"{concept} UNMAPPED in {latest}")
    try:
        facts = load_companyfacts(SecClient(ua), res.cik)
    except Exception as e:  # the facts are only evidence; never crash the run
        return bad, notes + [f"companyfacts unavailable ({type(e).__name__})"]
    units, dei = _first(res.driver_defaults.get("UNITS")), _basic_units(facts, latest)
    if isinstance(units, (int, float)) and units and dei:
        lo, hi = UNITS_RATIO
        if not lo <= units / dei <= hi:
            bad.append(f"UNITS {units:,.1f} MM vs dei shares outstanding {dei:,.1f} MM "
                       f"(ratio {units / dei:.2f}, outside {lo}-{hi}x)")
    concepts = {c.name: c for c in CONCEPTS}
    for name in SIGN_CONCEPTS:
        cv = res.mapping.get(name, {}).get(latest)
        if cv is None or not cv.value or not (cv.tag or "").startswith("us-gaap_"):
            continue
        fact = fy_values(facts, cv.tag.removeprefix("us-gaap_"), res.currency).get(latest)
        if not fact:
            continue
        expected = concepts[name].tag_sign(cv.tag) * fact   # income-positive tags: -1
        tol = max(1.0, 0.005 * abs(cv.value))
        if abs(cv.value + expected) <= tol < abs(cv.value - expected):
            bad.append(f"{name} {latest} mapped {cv.value:,.1f} but XBRL {cv.tag} = "
                       f"{fact:,.1f} (negated label carried into the model)")
    return bad, notes


def excel_check(path: Path) -> tuple[str, list[str]]:
    """Recalculate in Excel; return ("perp / exit" value per unit, error cells)."""
    import win32com.client as win32
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    try:
        wb = xl.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        try:
            xl.CalculateFull()
            errors = []
            for ws in wb.Worksheets:
                used = ws.UsedRange
                vals = used.Value
                vals = vals if isinstance(vals, tuple) else ((vals,),)
                for i, row in enumerate(vals):
                    for j, v in enumerate(row):
                        if isinstance(v, int) and v in XL_ERRORS:
                            errors.append(f"{ws.Name}!R{used.Row + i}C{used.Column + j}="
                                          f"{XL_ERRORS[v]}")
            dcf = wb.Worksheets("DCF")
            r = next(r for r in range(1, 200) if dcf.Cells(r, 2).Value == "Value per unit")

            def fmt(v):
                return f"{v:,.2f}" if isinstance(v, float) else str(v)
            return f"{fmt(dcf.Cells(r, 3).Value)} / {fmt(dcf.Cells(r, 4).Value)}", errors
        finally:
            wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


def run_one(ticker: str, ua: str, years: int, out_dir: Path, min_years: int,
            verbose: bool, cik: int | None = None, excel: bool = False) -> Row:
    row = Row(ticker)
    progress = (lambda m: print(f"    [{ticker}] {m}", flush=True)) if verbose else (lambda m: None)
    t0 = time.perf_counter()
    try:
        res = build(ticker, ua, n_years=years, progress=progress, cik_override=cik)
    except UserFacingError as e:
        row.secs = time.perf_counter() - t0
        row.note = str(e)
        if ticker not in EXPECT_40F:
            row.failures.append(f"UserFacingError: {e}")
        elif "40-F" not in str(e):
            row.failures.append(f"expected a 40-F message, got: {e}")
        return row
    except Exception as e:  # build() should never let one through
        row.secs = time.perf_counter() - t0
        row.failures.append(f"unexpected {type(e).__name__}: {e}")
        return row
    row.secs = time.perf_counter() - t0
    if ticker in EXPECT_40F:
        row.failures.append("expected a 40-F UserFacingError, but the build succeeded")

    fys = res.fy_labels
    row.n_years, row.currency = len(fys), res.currency
    row.years = f"{fys[0]}-{fys[-1]}" if fys else "-"
    missing = [f"{fy}:{st}" for fy in fys for st in STYPES
               if fy not in res.statements[st].sources]
    full = sum(all(fy in res.statements[st].sources for st in STYPES) for fy in fys)
    row.stmts = f"{full}/{len(fys)}"
    row.counts = summarize(res.checks)
    row.coverage = list(res.coverage)
    row.bal = res.bal_check_max
    if res.dropped_years:
        row.note = "dropped " + ",".join(res.dropped_years)

    if len(fys) < min_years:
        row.failures.append(f"only {len(fys)} fiscal years (< {min_years})")
    if missing:
        row.failures.append("missing statements: " + ", ".join(missing))
    if row.counts.get("FAIL", 0):
        fails = [f"{c.name} {c.fy}" for c in res.checks if c.status == "FAIL"]
        row.failures.append(f"{len(fails)} FAIL checks: " + "; ".join(fails[:8])
                            + (" ..." if len(fails) > 8 else ""))
    if row.bal is None or not row.bal < BAL_TOL:
        row.failures.append(f"max |BAL_CHECK| = {row.bal} (>= {BAL_TOL})")
    row.sanity, row.notes = sanity_checks(res, ua)
    row.failures.extend(f"sanity: {n}" for n in row.sanity)

    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"{ticker}.xlsx"
    path.write_bytes(res.xlsx)
    row.kb = len(res.xlsx) // 1024
    refs, load_err = scan_ref_errors(path)
    if load_err:
        row.failures.append(f"openpyxl.load_workbook failed: {load_err}")
    if refs:
        row.failures.append(f"#REF! in {len(refs)} cells: " + ", ".join(refs[:8]))
    if excel:
        try:
            row.vpu, xl_errors = excel_check(path)
        except Exception as e:  # Excel missing or COM failure: report, do not crash
            row.failures.append(f"Excel check failed: {type(e).__name__}: {e}")
        else:
            if xl_errors:
                row.failures.append(f"{len(xl_errors)} Excel error cells: "
                                    + ", ".join(xl_errors[:8]))
    return row


def line(r: Row) -> str:
    if not r.years:
        return f"{r.ticker:<5} {r.status:<4} {r.secs:6.1f}s  error: {r.note}"
    c = r.counts
    bal = "n/a" if r.bal is None else f"{r.bal:.1e}"
    return (f"{r.ticker:<5} {r.status:<4} {r.n_years:>2}y {r.years:<13} {r.currency:<3} "
            f"stmts {r.stmts:<5} P/W/F/S {c.get('PASS', 0)}/{c.get('WARN', 0)}/"
            f"{c.get('FAIL', 0)}/{c.get('SKIP', 0)}  cov {len(r.coverage)}  bal {bal}  "
            f"{r.secs:6.1f}s  {r.kb} KB" + (f"  value/unit {r.vpu}" if r.vpu else "")
            + (f"  ({r.note})" if r.note else ""))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("tickers", nargs="+")
    ap.add_argument("--ua", required=True, help='SEC User-Agent, "Name email@domain"')
    ap.add_argument("--years", type=int, default=10, help="fiscal years to request")
    ap.add_argument("--min-years", type=int, default=5, help="fewest years that pass")
    ap.add_argument("--out", type=Path, default=ROOT / "out")
    ap.add_argument("-v", "--verbose", action="store_true", help="echo pipeline progress")
    ap.add_argument("--cik", action="append", default=[], metavar="TICKER=CIK",
                    help="build TICKER with a CIK override (repeatable)")
    ap.add_argument("--excel", action="store_true",
                    help="recalculate in Excel (COM): assert no error cells, print value/unit")
    args = ap.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    logging.basicConfig(level=logging.WARNING, stream=sys.stderr)

    ciks = {}
    for spec in args.cik:
        t, _, c = spec.partition("=")
        ciks[t.strip().upper()] = int(c)
    rows = []
    for t in [t.strip().upper() for t in args.tickers]:
        r = run_one(t, args.ua, args.years, args.out, args.min_years, args.verbose,
                    ciks.get(t), args.excel)
        expected = KNOWN_FAILURES.get(t) if t not in ciks else None
        if expected and r.failures and all(expected in f for f in r.failures):
            r.known = True
        elif expected and not r.failures:
            r.failures.append(f"documented failure ({expected!r}) no longer occurs: "
                              "update KNOWN_FAILURES and the acceptance doc")
        rows.append(r)
        print(line(r), flush=True)
        for w in r.coverage:
            print(f"      coverage: {w}", flush=True)
        for n in r.notes:
            print(f"      NOTE: {n}", flush=True)
        for f in r.failures:
            print(f"      {'KNOWN' if r.known else 'ASSERT'}: {f}", flush=True)

    print("\n=== Summary ===")
    for r in rows:
        print(line(r))
    bad = [r for r in rows if r.failures and not r.known]
    known = [r.ticker for r in rows if r.known]
    print(f"\n{len(rows) - len(bad) - len(known)}/{len(rows)} tickers passed"
          + (f"; known (documented) failures: {', '.join(known)}" if known else "")
          + (f"; failed: {', '.join(r.ticker for r in bad)}" if bad else ""))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
