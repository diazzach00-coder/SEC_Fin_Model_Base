"""Run one ticker end to end: EDGAR -> statements -> mapping -> checks -> model -> .xlsx."""
import logging
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import requests

from .checks import API_CONCEPTS, run_checks, summarize
from .companyfacts import (dei_shares_outstanding, fill_from_facts, fy_values,
                           load_companyfacts, unit_hint_from_facts)
from .edgar import (choose_source_filing, company_for_cik, list_annual_filings,
                    resolve_ticker, select_filings_by_fy)
from .forecast import compute, driver_defaults, scalar_defaults
from .http import SecClient
from .mapping import coverage_warnings, normalize_expense_signs, resolve
from .merge import merge_statements
from .model_spec import build_spec
from .rfile import parse_rfile
from .types import CheckResult, Filing, MarketInputs, MergedStatement, UserFacingError
from .workbook.build import build_workbook, fye_text

log = logging.getLogger(__name__)

STYPES = ("IS", "BS", "CF")
TOO_FEW_YEARS = "Fewer than two fiscal years with full XBRL statements found."
FILL_TAGS = {
    "DilutedUnits": ["WeightedAverageLimitedPartnershipUnitsOutstandingDiluted",
                     "WeightedAverageNumberOfDilutedSharesOutstanding"],
    "UnitsOutstanding": ["dei:EntityCommonStockSharesOutstanding", "CommonStockSharesOutstanding",
                         "LimitedPartnersCapitalAccountUnitsOutstanding"],
}


@dataclass
class BuildResult:
    company: str
    ticker: str
    cik: int
    currency: str
    fy_labels: list[str]
    fcst_labels: list[str]
    statements: dict[str, MergedStatement]
    mapping: dict
    checks: list[CheckResult]
    coverage: list[str]
    forecast: dict
    xlsx: bytes
    log: list[str]
    bal_check_max: float = 0.0
    driver_defaults: dict = field(default_factory=dict)
    scalar_defaults: dict = field(default_factory=dict)
    da_basis: str = ""
    dropped_years: list[str] = field(default_factory=list)
    company_fye: str = ""


def build(ticker: str, user_agent: str, n_years: int = 10, n_fcst: int = 5,
          market: MarketInputs = MarketInputs(), progress: Callable[[str], None] = print,
          cache_dir=Path(".cache"), *, client=None,
          cik_override: int | None = None) -> BuildResult:
    """Build the workbook for ``ticker``. Every failure reaches the caller as a
    ``UserFacingError``; unexpected ones are logged with their traceback first.

    ``client`` (anything with ``get_json``/``get_text``) replaces the SEC client, in which
    case ``user_agent`` and ``cache_dir`` are unused. ``cik_override`` (a positive CIK)
    skips ticker resolution, e.g. for a successor holding company whose ticker maps to a
    CIK with no 10-K yet (XOM in 2026: history is under the predecessor CIK 34088).
    """
    t = ticker.strip().upper()
    try:
        if not t:
            raise UserFacingError("Enter a ticker.")
        if client is None:
            client = SecClient(user_agent, cache_dir)
        return _Run(t, client, progress, cik_override or None).run(n_years, n_fcst, market)
    except UserFacingError:
        raise
    except Exception as e:
        log.exception("Build failed for %s", t)
        raise UserFacingError(f"Build failed for {t}: {type(e).__name__}. See server log.") from e


class _Run:
    def __init__(self, ticker: str, client, progress: Callable[[str], None],
                 cik_override: int | None = None):
        self.ticker, self.client, self.progress = ticker, client, progress
        self.cik_override = cik_override
        self.log: list[str] = []
        self.no_filing: set[str] = set()   # FYs none of whose filings has all three statements

    def say(self, msg: str) -> None:
        self.log.append(msg)
        self.progress(msg)

    def run(self, n_years: int, n_fcst: int, market: MarketInputs) -> BuildResult:
        t, client = self.ticker, self.client
        if self.cik_override:
            cik = int(self.cik_override)
            company = company_for_cik(client, cik)
            self.say(f"Using CIK override {cik} -> {company} for {t}")
        else:
            cik, company = resolve_ticker(client, t)
            self.say(f"Resolved {t} -> {company} (CIK {cik})")

        try:
            filings = list_annual_filings(client, cik)
        except UserFacingError as e:   # "CIK 123 files Form 40-F ..." -> "OKE files Form 40-F ..."
            raise UserFacingError(re.sub(r"^CIK \d+ ", f"{t} ", str(e))) from None
        if not filings:
            raise UserFacingError(f"No 10-K filings found for {t}.")
        by_fy = select_filings_by_fy(filings, n_years)
        self.say(f"Found {len(filings)} annual filings; checking {len(by_fy)} fiscal years")

        facts = self._companyfacts(cik)
        parsed = self._parse_filings(by_fy, unit_hint_from_facts(facts))
        fy_labels, fy_ends, dropped = self._fiscal_years(parsed, by_fy, n_years)

        used = [(f, s) for f, s in parsed.values() if f.fy_label in fy_ends]
        statements = {st: merge_statements([(f, s[st]) for f, s in used], fy_labels, fy_ends)
                      for st in STYPES}
        currency = statements["IS"].currency

        mapping = normalize_expense_signs(
            resolve(statements), lambda tag: fy_values(facts, tag, currency))
        for concept, tags in FILL_TAGS.items():
            mapping[concept] = fill_from_facts(mapping, facts, concept, tags, "shares")
        coverage = coverage_warnings(mapping)
        self.say(f"Mapping: {len(coverage)} coverage warning{'' if len(coverage) == 1 else 's'}")

        checks = run_checks(statements, mapping, _facts_values(mapping, facts, currency))
        n = summarize(checks)
        self.say(f"Checks: {n['PASS']} PASS, {n['WARN']} WARN, {n['FAIL']} FAIL, {n['SKIP']} SKIP")

        history = {c: {fy: cv.value for fy, cv in years.items()} for c, years in mapping.items()}
        latest_ppe = mapping.get("NetPPE", {}).get(fy_labels[-1])
        da_basis = ("ppe" if latest_ppe is not None and latest_ppe.status != "UNMAPPED"
                    and latest_ppe.value is not None else "revenue")
        spec = build_spec(da_basis)
        drivers = driver_defaults(spec, history, fy_labels)
        scalars = scalar_defaults(spec, history, fy_labels)
        last_year = fy_ends[fy_labels[-1]].year
        fcst_labels = [f"FY{last_year + i}E" for i in range(1, n_fcst + 1)]
        forecast = compute(spec, history, fy_labels, fcst_labels)
        bal = max((abs(forecast["BAL_CHECK"].get(fy) or 0.0) for fy in fcst_labels), default=0.0)
        self.say(f"Forecast {'balances' if bal < 1e-6 else 'does NOT balance'} "
                 f"(max |check| {bal:.4f})")

        xlsx = build_workbook(
            company=company, ticker=t, cik=cik, currency=currency, fy_labels=fy_labels,
            fcst_labels=fcst_labels, fy_ends=fy_ends, statements=statements, mapping=mapping,
            checks=checks, spec=spec, drivers=drivers, scalars=scalars, market=market,
            basic_units=_basic_units(facts, fy_labels[-1]),
            filings_per_fy=_filings_per_fy(statements, fy_labels))
        self.say(f"Workbook built ({len(xlsx) // 1024} KB)")

        return BuildResult(
            company=company, ticker=t, cik=cik, currency=currency, fy_labels=fy_labels,
            fcst_labels=fcst_labels, statements=statements, mapping=mapping, checks=checks,
            coverage=coverage, forecast=forecast, xlsx=xlsx, log=self.log,
            bal_check_max=bal, driver_defaults=drivers, scalar_defaults=scalars,
            da_basis=da_basis, dropped_years=dropped, company_fye=fye_text(fy_ends[fy_labels[-1]]))

    # -- steps ----------------------------------------------------------------
    def _companyfacts(self, cik: int) -> dict:
        try:
            return load_companyfacts(self.client, cik)
        except requests.HTTPError:
            self.say("Companyfacts API unavailable; API cross-checks will be skipped")
            return {}

    def _parse_filings(self, by_fy: dict[str, list[Filing]], unit_hint) -> dict:
        """accession -> (Filing, {stype: ParsedStatement}) for each FY's chosen filing;
        a filing chosen for several years is parsed once."""
        parsed: dict[str, tuple[Filing, dict]] = {}
        for fy, candidates in by_fy.items():
            chosen = choose_source_filing(self.client, candidates)
            if chosen is None:           # reported by _fiscal_years (it may still be covered)
                self.no_filing.add(fy)
                continue
            filing, located = chosen
            self.say(f"{fy}: {filing.form} {filing.accession} -> " + ", ".join(
                f"{st} {located[st].html_file.removesuffix('.htm')}" for st in STYPES))
            if filing.accession in parsed:
                continue
            stmts = {}
            for st in STYPES:
                url = filing.folder_url + located[st].html_file
                stmts[st] = parse_rfile(self.client.get_text(url), st, url, unit_hint)
            parsed[filing.accession] = (filing, stmts)
        self.say(f"Parsed {len(parsed)} filing{'' if len(parsed) == 1 else 's'}")
        return parsed

    def _fiscal_years(self, parsed: dict, by_fy: dict, n_years: int):
        """Years with a full-year IS and CF column and a BS column in the chosen filings
        (a filing's comparative columns can cover a year whose own filing lacks statements),
        newest ``n_years``. Returns (fy_labels, fy_ends, dropped_years)."""
        ends = {st: set() for st in STYPES}
        for _, stmts in parsed.values():
            for st, ps in stmts.items():
                want = None if st == "BS" else 12
                ends[st] |= {c.end for c in ps.columns if c.months == want}
        complete = ends["IS"] & ends["BS"] & ends["CF"]
        qualified: dict[str, date] = {}
        for end in sorted(complete):
            qualified[f"FY{end.year}"] = end
        for filing, _ in parsed.values():   # a filing's own year end wins a label clash
            if filing.report_date in complete:
                qualified[filing.fy_label] = filing.report_date

        kept = sorted(qualified)[-n_years:] if n_years > 0 else []
        seen = set(by_fy) | {f"FY{e.year}" for st in STYPES for e in ends[st]}
        floor = kept[0] if kept and len(kept) == n_years else min(seen, default="")
        dropped = sorted(fy for fy in seen if fy not in qualified and fy >= floor)
        for fy in sorted(set(dropped) | (self.no_filing & set(kept))):
            self.say(f"{fy}: no filing with all three statements; " + (
                "dropped" if fy in dropped else "covered by a later filing's comparative columns"))
        if len(kept) < 2:
            raise UserFacingError(TOO_FEW_YEARS)
        self.say(f"Using {len(kept)} fiscal years: {kept[0]}-{kept[-1]}")
        return kept, {fy: qualified[fy] for fy in kept}, dropped


def _facts_values(mapping: dict, facts: dict, currency: str) -> dict[str, dict[str, float]]:
    """companyfacts values for the API cross-check concepts: each year is read with the
    us-gaap tag the mapping chose for that year (a filer's tag can change, e.g. TRGP's
    FY2016 revenue is ``RevenueFromContractWithCustomer...`` but ``Revenues`` later).
    Custom, SUM(...) and companyfacts tags are skipped."""
    out: dict[str, dict[str, float]] = {}
    by_tag: dict[str, dict[str, float]] = {}
    for concept in API_CONCEPTS:
        for fy, cv in (mapping.get(concept) or {}).items():
            tag = cv.tag if cv is not None else None
            if not tag or not re.fullmatch(r"us-gaap_\w+", tag):
                continue
            if tag not in by_tag:
                by_tag[tag] = fy_values(facts, tag.removeprefix("us-gaap_"), currency)
            if fy in by_tag[tag]:
                out.setdefault(concept, {})[fy] = by_tag[tag][fy]
    return out


def _basic_units(facts: dict, last_fy: str) -> float | None:
    """Latest cover-page shares outstanding (MM) for a year no later than ``last_fy``."""
    years = [fy for fy in dei_shares_outstanding(facts).items() if fy[0] <= last_fy]
    return max(years)[1] if years else None


def _filings_per_fy(statements: dict[str, MergedStatement], fy_labels: list[str]) -> list[tuple]:
    """(fy, form, accession, filing_date) of the filing that supplied each year."""
    out = []
    for fy in fy_labels:
        src = next((statements[st].sources[fy] for st in STYPES if fy in statements[st].sources),
                   None)
        if src is not None:
            out.append((fy, src.form, src.accession, src.filing_date))
    return out
