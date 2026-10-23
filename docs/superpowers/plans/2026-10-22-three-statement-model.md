# SEC 3-Statement Model Builder Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Type a ticker, get an Excel workbook with the as-reported 10-K IS/BS/CF history keyed by XBRL tag, tie-out checks, a linked 5-year 3-statement forecast driven by an asset block, and WACC + DCF (perpetuity and EBITDA exit multiple) with live formulas.

**Architecture:** A `sec_model` package: `edgar` finds filings and statement R-files via FilingSummary.xml; `rfile` parses one R-file into tag-keyed rows and dated columns; `merge` stacks filings into one statement per type (latest filing wins per year); `mapping` resolves model concepts to rows; `checks` runs tie-outs; `formula` is a tiny expression AST rendered both to Excel formulas and evaluated in Python, so the Model/DCF sheets and the regression script share one definition; `workbook` writes the sheets with openpyxl; `app.py` is the Streamlit UI.

**Tech Stack:** Python 3.11+, requests, beautifulsoup4, pandas, openpyxl, streamlit, pytest (dev). No other runtime deps.

**Spec:** `docs/superpowers/specs/2026-10-22-three-statement-model-design.md`

## Global Constraints

- Runtime deps are exactly `streamlit requests pandas openpyxl beautifulsoup4` (requirements.txt). Dev deps go in `requirements-dev.txt` (`pytest`).
- SEC requests: User-Agent string from the user (must contain `@`), `Accept-Encoding: gzip, deflate`, ≥0.11 s between requests, 3 retries with backoff 1/2/4 s, 30 s timeout. Every response cached on disk under `.cache/` (gitignored), keyed by sha1(url), 7-day expiry.
- All stored values: dollars in **millions**, shares/units in **millions**, per-share/per-unit unscaled. Currency detected, never converted.
- Excel: Calibri; inputs blue font + yellow fill `FFF2CC`; formulas black; key outputs green fill `D9EAD3`; history columns start at column D; forecast columns fill `EBF3FB`; no Excel Data Tables; no `#REF!`; no macros.
- Never raise to the UI. `pipeline.build()` returns a `BuildResult` or raises `UserFacingError(message)`; the app shows `message` only.
- Supported: 10-K / 10-K/A / 10-KT filers with XBRL. 40-F/20-F → `UserFacingError("TICKER files Form 40-F (no US-GAAP XBRL statements); not supported.")`.
- Fiscal-year label = `FY` + year of the fiscal-year end date.
- Forecast years default 5, history default 10 (bounded by what exists).

## Review Focus

1. A 10-K whose IS R-file has a blank first value column plus `[1]` footnote markers (ET FY2025) → columns are keyed by period date, so FY values land in the right year and `[1]` never becomes a value. Test in Task 4.
2. A company that changes a tag between years (EPD custom `epd_DepreciationAndAccretion` vs standard D&A) → the D&A concept resolves per year, no `UNMAPPED` year, and the Mapping sheet shows the different tags. Test in Task 6.
3. A 10-K/A that contains no financial statements (Part III amendment) listed after the original → the original 10-K supplies the year; the build does not report "statement missing". Test in Task 3.
4. A company with no `Liabilities` and no `TotalCurrentLiabilities` tag (EPD, OKE) → the Model's "Other liabilities (derived)" absorbs it and the historical balance check still passes. Test in Task 9.
5. A forecast year where cash before revolver drops below minimum cash → revolver draws exactly the shortfall, ending cash = minimum cash, BS balances. Test in Task 9.

## Shared types (`sec_model/types.py`, created in Task 1)

```python
@dataclass(frozen=True)
class Filing:
    accession: str          # "0001039684-26-000012"
    form: str               # "10-K" | "10-K/A" | "10-KT"
    filing_date: date
    report_date: date       # fiscal year end
    primary_doc: str
    @property
    def fy_label(self) -> str: ...        # "FY2025"
    @property
    def folder_url(self) -> str: ...      # https://www.sec.gov/Archives/edgar/data/{int cik}/{acc no dashes}/

@dataclass(frozen=True)
class Column:
    end: date
    months: int | None      # None for instant (BS)
    label: str              # header text as shown

@dataclass
class Row:
    key: str                # tag, or f"{tag}|{label}" for dimensional duplicates
    tag: str                # "us-gaap_ProfitLoss", "epd_DepreciationAndAccretion"
    label: str
    indent: int
    is_abstract: bool
    unit: str               # "money" | "shares" | "perunit" | "other"
    values: list[float | None]   # one per Column

@dataclass
class ParsedStatement:
    stype: str              # "IS" | "BS" | "CF"
    title: str
    currency: str           # "USD" | "CAD" ...
    money_scale: float      # multiplier to millions: 1 for "$ in Millions", 0.001 for thousands
    share_scale: float
    columns: list[Column]
    rows: list[Row]
    source_url: str

@dataclass
class SourceRec:
    accession: str; form: str; filing_date: date; r_url: str; scale_text: str

@dataclass
class MergedRow:
    key: str; tag: str; label: str; indent: int; is_abstract: bool; unit: str
    values: dict[str, float | None]     # fy_label -> value
    still_reported: bool

@dataclass
class MergedStatement:
    stype: str
    currency: str
    fy_labels: list[str]                # oldest -> newest
    fy_ends: dict[str, date]
    rows: list[MergedRow]
    sources: dict[str, SourceRec]       # fy_label -> source

@dataclass
class ConceptValue:
    concept: str; fy: str; value: float | None
    row_key: str | None; tag: str | None; label: str | None
    status: str                         # "OK" | "UNMAPPED" | "DERIVED"

@dataclass
class CheckResult:
    name: str; fy: str; status: str     # "PASS" | "WARN" | "FAIL" | "SKIP"
    lhs: float | None; rhs: float | None; tolerance: float; note: str

@dataclass
class MarketInputs:
    price: float | None = None
    beta: float = 1.0
    risk_free: float = 0.0425
    mrp: float = 0.055

class UserFacingError(Exception): ...
```

---

### Task 1: HTTP layer with disk cache

**Files:**
- Create: `sec_model/__init__.py`, `sec_model/types.py` (as above), `sec_model/http.py`, `requirements-dev.txt`, `.gitignore` (add `.cache/`, `__pycache__/`, `*.xlsx` under `out/`)
- Test: `tests/test_http.py`

**Interfaces:**
- Produces: `class SecClient: __init__(self, user_agent: str, cache_dir: Path = Path(".cache"), min_interval: float = 0.11)`; `get_json(url) -> dict`; `get_text(url) -> str`; `get_bytes(url) -> bytes`. Raises `UserFacingError("SEC User-Agent must include an email address")` if `"@" not in user_agent`. 404 raises `requests.HTTPError`; the caller decides.

- [ ] **Step 1: Write failing tests**

```python
def test_rejects_user_agent_without_email(tmp_path):
    with pytest.raises(UserFacingError):
        SecClient("no email here", cache_dir=tmp_path)

def test_second_call_hits_cache(tmp_path, monkeypatch):
    calls = []
    def fake_get(url, headers, timeout):
        calls.append(url); return FakeResp(200, b'{"a":1}')
    monkeypatch.setattr("sec_model.http.requests.get", fake_get)
    c = SecClient("Test test@example.com", cache_dir=tmp_path)
    assert c.get_json("https://data.sec.gov/x.json") == {"a": 1}
    assert c.get_json("https://data.sec.gov/x.json") == {"a": 1}
    assert len(calls) == 1

def test_host_header_matches_url(tmp_path, monkeypatch):
    seen = {}
    monkeypatch.setattr("sec_model.http.requests.get",
                        lambda url, headers, timeout: seen.update(headers) or FakeResp(200, b"ok"))
    SecClient("Test test@example.com", cache_dir=tmp_path).get_text("https://www.sec.gov/y.htm")
    assert seen["Host"] == "www.sec.gov" and "@" in seen["User-Agent"]

def test_retries_then_raises(tmp_path, monkeypatch):
    monkeypatch.setattr("sec_model.http.time.sleep", lambda s: None)
    monkeypatch.setattr("sec_model.http.requests.get", lambda *a, **k: (_ for _ in ()).throw(requests.ConnectionError()))
    with pytest.raises(requests.ConnectionError):
        SecClient("Test test@example.com", cache_dir=tmp_path).get_text("https://www.sec.gov/z")
```

- [ ] **Step 2: Run tests, expect ImportError**

Run: `python -m pytest tests/test_http.py -v` — Expected: FAIL (module not found)

- [ ] **Step 3: Implement `sec_model/types.py` and `sec_model/http.py`**

Cache file: `cache_dir / sha1(url).hexdigest()` with a sidecar `.meta` JSON holding `fetched_at`; expired after 7 days. Do not cache non-200.

- [ ] **Step 4: Run tests, expect PASS**

- [ ] **Step 5: Commit** — `git add sec_model tests requirements-dev.txt .gitignore && git commit -m "feat: SEC HTTP client with disk cache and shared types"`

---

### Task 2: Ticker lookup and 10-K filing selection

**Files:**
- Create: `sec_model/edgar.py`
- Test: `tests/test_edgar.py`

**Interfaces:**
- Consumes: `SecClient`
- Produces:
  - `resolve_ticker(client, ticker: str) -> tuple[int, str]` → `(cik, company_name)`; tries `company_tickers.json`, then `company_tickers_exchange.json` (its `data` rows: `[cik, name, ticker, exchange]`), matching case-insensitively with `.`/`-` treated as equal. Raises `UserFacingError(f"Ticker '{t}' not found in SEC company list.")`.
  - `list_annual_filings(client, cik: int) -> list[Filing]` → all `10-K`, `10-K/A`, `10-KT` from `filings.recent` **and** every `filings.files[*].name` page (`https://data.sec.gov/submissions/{name}`), sorted by `report_date` then `filing_date`. If the only annual forms are `40-F`/`20-F`, raise the 40-F `UserFacingError` from Global Constraints (message uses the form actually found).
  - `select_filings_by_fy(filings: list[Filing], n_years: int) -> dict[str, list[Filing]]` → `fy_label -> filings for that FY, newest filing first`, limited to the `n_years` most recent fiscal years. (Which one supplies statements is decided in Task 3 because it needs FilingSummary.)

- [ ] **Step 1: Write failing tests** (use small hand-written submissions dicts via a `FakeClient` whose `get_json` returns from a dict of url→payload)

```python
def test_resolve_ticker_case_and_class_suffix():
    # payload: {"0": {"cik_str": 1067983, "ticker": "BRK-B", "title": "BERKSHIRE..."}}
    assert resolve_ticker(fake, "brk.b") == (1067983, "BERKSHIRE HATHAWAY INC")

def test_list_annual_filings_reads_paginated_files():
    # recent has FY2025 10-K; files[0] page has FY2015 10-K
    fs = list_annual_filings(fake, 1039684)
    assert [f.fy_label for f in fs] == ["FY2015", "FY2025"]

def test_40f_filer_raises():
    with pytest.raises(UserFacingError, match="40-F"):
        list_annual_filings(fake_trp, 1232384)

def test_select_filings_by_fy_groups_amendments_newest_first():
    sel = select_filings_by_fy(filings, n_years=2)
    assert list(sel) == ["FY2023", "FY2024"]
    assert [f.form for f in sel["FY2024"]] == ["10-K/A", "10-K"]
```

- [ ] **Step 2: Run, expect FAIL** — `python -m pytest tests/test_edgar.py -v`
- [ ] **Step 3: Implement the three functions in `sec_model/edgar.py`**
- [ ] **Step 4: Run, expect PASS**
- [ ] **Step 5: Commit** — `git commit -am "feat: ticker resolution and annual filing selection"`

---

### Task 3: Find the three statements in a filing (FilingSummary.xml)

**Files:**
- Modify: `sec_model/edgar.py`
- Create: `scripts/fetch_fixtures.py`
- Test: `tests/test_edgar.py`, fixtures `tests/fixtures/<TICKER>/FilingSummary.xml`

**Interfaces:**
- Produces:
  - `@dataclass ReportEntry: short_name: str; long_name: str; category: str; html_file: str`
  - `parse_filing_summary(xml: str) -> list[ReportEntry]`
  - `pick_statement_reports(reports: list[ReportEntry]) -> dict[str, list[ReportEntry]]` → `{"IS": [...candidates best first], "BS": [...], "CF": [...]}`. Only `category == "Statements"`. Exclude names containing (case-insensitive): `parenthetical`, `comprehensive`, `equity`, `partners' capital`, `details`, `stockholders`, `changes in`. Score with the existing `_PRECISE`/`_BROAD` keyword lists from `streamlit_app.py:106-117` (copy them, extend IS precise with `"statements of consolidated operations"`, `"statements of earnings"`; BS with `"financial position"`; CF with `"statements of consolidated cash"`). Candidates with score 0 are dropped.
  - `locate_statements(client, filing: Filing) -> dict[str, ReportEntry | None]` → fetches `FilingSummary.xml` (404 → all `None`), picks the **first candidate whose R-file contains the anchor tag** (`us-gaap_Revenues|us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax|us-gaap_CostsAndExpenses|us-gaap_InterestAndDividendIncomeOperating` for IS; `us-gaap_Assets` for BS; `us-gaap_NetCashProvidedByUsedInOperatingActivities` for CF — substring search on the HTML, which Task 4 parses properly). Caches the HTML via `SecClient` so Task 4 does not refetch.
  - `choose_source_filing(client, candidates: list[Filing]) -> tuple[Filing, dict[str, ReportEntry | None]] | None` → first filing (newest first) for which `locate_statements` finds at least IS and BS and CF; else `None`.

- [ ] **Step 1: Write `scripts/fetch_fixtures.py`**: `python scripts/fetch_fixtures.py OKE EPD ET KNTK ORCL JPM --ua "Name email@x.com"` → for each ticker, latest 10-K: saves `FilingSummary.xml`, and the chosen `IS.htm`, `BS.htm`, `CF.htm` to `tests/fixtures/<TICKER>/`, and prints each statement's title line (so the implementer can confirm KNTK is `$ in Thousands` and ORCL is a May FYE; if KNTK is in millions, use `MPLX` or another filer whose title says Thousands and note which in the fixture folder's `README.txt`).
- [ ] **Step 2: Run it** and commit the fixtures (they are a few hundred KB).
- [ ] **Step 3: Write failing tests**

```python
def test_pick_statement_reports_oke():
    reps = parse_filing_summary(FIX("OKE/FilingSummary.xml"))
    best = {k: v[0].html_file for k, v in pick_statement_reports(reps).items()}
    assert best == {"IS": "R3.htm", "BS": "R7.htm", "CF": "R9.htm"}

def test_pick_statement_reports_epd_bs_before_is():
    # EPD lists BS as R3, IS ("Statements of Consolidated Operations") R5, CF R7
    ...assert best == {"IS": "R5.htm", "BS": "R3.htm", "CF": "R7.htm"}

def test_pick_excludes_parenthetical_and_comprehensive():
    reps = [ReportEntry("CONSOLIDATED STATEMENTS OF INCOME (Parenthetical)", "", "Statements", "R4.htm"),
            ReportEntry("CONSOLIDATED STATEMENTS OF COMPREHENSIVE INCOME", "", "Statements", "R5.htm"),
            ReportEntry("CONSOLIDATED STATEMENTS OF INCOME", "", "Statements", "R3.htm")]
    assert [r.html_file for r in pick_statement_reports(reps)["IS"]] == ["R3.htm"]

def test_choose_source_filing_skips_part3_amendment():
    # fake client: 10-K/A FilingSummary has only Cover; original 10-K has all three
    f, located = choose_source_filing(fake, [amend, original])
    assert f is original and all(located.values())
```

- [ ] **Step 4: Run, expect FAIL** → **Step 5: Implement** → **Step 6: Run, expect PASS**
- [ ] **Step 7: Commit** — `git add -A scripts tests sec_model && git commit -m "feat: locate IS/BS/CF R-files via FilingSummary; add fixtures"`

---

### Task 4: R-file parser

**Files:**
- Create: `sec_model/rfile.py`
- Test: `tests/test_rfile.py`

**Interfaces:**
- Produces: `parse_rfile(html: str, stype: str, source_url: str, unit_hint: Callable[[str], str] | None = None) -> ParsedStatement`
  - `unit_hint(tag) -> "money"|"shares"|"perunit"|"other"|None` is supplied by the pipeline from companyfacts units (Task 7); when it returns `None` (custom tag) fall back to label rules: `per (share|unit)` → `perunit`; `(shares|units)` + (`weighted|outstanding|issued`) → `shares`; else `money`.
  - Header: title cell text; `currency` = regex `([A-Z]{3}) \(\$\)` → group 1, default `USD`; `money_scale` from `\$ in (Millions|Thousands|Billions)` → 1 / 0.001 / 1000, default 1e-6 (raw dollars); `share_scale` from `shares in (Millions|Thousands)` → 1 / 0.001, default 1e-6.
  - Columns: duration statements have a row with `N Months Ended` spanning columns and a row of dates; instant statements have one date row. Produce one `Column(end, months, label)` per data column. Parse dates like `Dec. 31, 2025` / `Dec. 31, 2024`.
  - Rows: `<a ... onclick="...defref_(\w+)...">` gives `tag`; `is_abstract = tag.endswith("Abstract")`; indent = round(`padding-left` em from the first cell's `style`, else count of nested `div`/`span` with `pl` classes, else 0); label = cell text with `\[\d+\]` and surrounding whitespace removed. Values: strip `$`, `,`, whitespace, `[n]` markers; `(x)` → negative; `—`/`–`/`-`/`` → `None`; apply scale by unit (`perunit` unscaled). Rows whose tag is already used in this statement get `key = f"{tag}|{label}"`.
  - Rows that are not statement data (e.g. the ET `[1]`-only cells) must yield `None`, never raise.

- [ ] **Step 1: Write failing tests against fixtures**

```python
def test_oke_cf_columns_and_scale():
    s = parse_rfile(FIX("OKE/CF.htm"), "CF", "u")
    assert s.currency == "USD" and s.money_scale == 1
    assert [c.end.year for c in s.columns] == [2025, 2024, 2023] and s.columns[0].months == 12

def test_oke_cf_first_rows():
    s = parse_rfile(FIX("OKE/CF.htm"), "CF", "u")
    ni = next(r for r in s.rows if r.tag == "us-gaap_ProfitLoss")
    assert ni.values == [3462.0, 3112.0, 2659.0] and ni.indent >= 1
    assert s.rows[0].is_abstract

def test_et_blank_first_column_and_footnote():
    s = parse_rfile(FIX("ET/CF.htm"), "CF", "u")
    ni = next(r for r in s.rows if r.tag == "us-gaap_ProfitLoss")
    assert ni.values[0] is None and ni.values[1] == 5708.0
    eq = next(r for r in s.rows if r.tag == "us-gaap_IncomeLossFromEquityMethodInvestments")
    assert eq.values[0] is None   # "[1]" is not a number

def test_epd_custom_tag_kept():
    s = parse_rfile(FIX("EPD/CF.htm"), "CF", "u")
    assert any(r.tag == "epd_DepreciationAndAccretion" and r.values[0] == 2107.0 for r in s.rows)

def test_thousands_filer_scaled_to_millions():
    s = parse_rfile(FIX("<THOUSANDS_TICKER>/BS.htm"), "BS", "u")
    assets = next(r for r in s.rows if r.tag == "us-gaap_Assets")
    assert s.money_scale == 0.001 and assets.values[0] < 1e6   # i.e. in $MM

def test_bs_is_instant():
    s = parse_rfile(FIX("OKE/BS.htm"), "BS", "u")
    assert all(c.months is None for c in s.columns) and len(s.columns) == 2

def test_per_share_not_scaled():
    s = parse_rfile(FIX("OKE/IS.htm"), "IS", "u")
    eps = next(r for r in s.rows if "EarningsPerShareDiluted" in r.tag)
    assert eps.unit == "perunit" and 0 < eps.values[0] < 100

def test_non_december_fye():
    s = parse_rfile(FIX("ORCL/IS.htm"), "IS", "u")
    assert s.columns[0].end.month == 5

def test_bank_layout_parses():
    s = parse_rfile(FIX("JPM/IS.htm"), "IS", "u")
    assert len([r for r in s.rows if not r.is_abstract]) >= 20
    assert not any(r.tag == "us-gaap_OperatingIncomeLoss" for r in s.rows)   # parser assumes nothing about layout

def test_duplicate_tag_gets_label_key():
    html = TWO_REVENUE_ROWS_HTML   # minimal table with two us-gaap_Revenues rows, labels "Product" and "Service"
    s = parse_rfile(html, "IS", "u")
    assert [r.key for r in s.rows] == ["us-gaap_Revenues", "us-gaap_Revenues|Service"]
```

- [ ] **Step 2: Run, expect FAIL** → **Step 3: Implement `sec_model/rfile.py`** → **Step 4: Run, expect PASS**
- [ ] **Step 5: Commit** — `git commit -am "feat: R-file parser keyed by XBRL tag and period date"`

---

### Task 5: Merge filings into one statement per type

**Files:**
- Create: `sec_model/merge.py`
- Test: `tests/test_merge.py`

**Interfaces:**
- Consumes: `ParsedStatement`, `Filing`
- Produces: `merge_statements(parsed: list[tuple[Filing, ParsedStatement]], fy_labels: list[str], fy_ends: dict[str, date]) -> MergedStatement`
  - A parsed column matches a fiscal year when `column.end == fy_ends[fy]` and (`months is None` or `months == 12`). Other columns are ignored.
  - For each `(fy, row key)`, value = from the parsed statement of the **newest filing** (by `filing_date`) that has that column.
  - Row order: iterate the newest filing's rows in order; then for each older filing (newest→oldest), any key not yet placed is inserted immediately after the key that precedes it in that older filing (or appended to the end if its predecessor is also absent), with `still_reported=False`.
  - `label`, `indent`, `is_abstract`, `unit` from the newest filing containing the key.
  - `sources[fy] = SourceRec(...)` of the filing that supplied that year's values.
  - `currency` from the newest filing; if filings disagree, raise `UserFacingError("Currency changes across filings (USD→CAD); not supported.")`.

- [ ] **Step 1: Write failing tests** (build `ParsedStatement`s by hand with 2–3 rows)

```python
def test_restated_value_from_newest_filing_wins(): ...   # FY2023 appears in 2024 and 2025 filings; 2025's value used
def test_label_change_same_tag_is_one_row(): ...         # "Revenues" (2024) vs "Total revenues" (2025) -> one row, label "Total revenues"
def test_dropped_row_inserted_after_predecessor_and_flagged(): ...
def test_non_12_month_column_ignored(): ...              # a 1-month stub column with end == fy end must not supply the value
def test_sources_record_filing_per_year(): ...
```

- [ ] **Step 2–4: FAIL → implement → PASS**
- [ ] **Step 5: Commit** — `git commit -am "feat: merge statements across filings, newest filing wins"`

---

### Task 6: Concept mapping

**Files:**
- Create: `sec_model/concepts.py` (the table), `sec_model/mapping.py`
- Test: `tests/test_mapping.py`

**Interfaces:**
- Produces:
  - `concepts.py`: `CONCEPTS: list[Concept]` with `@dataclass Concept: name: str; stype: str; tags: tuple[str, ...]; label_re: str | None; sign: int = 1; sum_tags: tuple[str, ...] = ()`. Tags are written without prefix (`"Revenues"`) and match `us-gaap_<tag>`; `dei:` prefixed tags are not in statements (handled in Task 7). `sum_tags` means the concept is the sum of several rows (used by `TotalDebt`). Include every concept in spec §5.6 plus: `OperatingIncome`, `TotalCostsAndExpenses`, `Impairments` (IS), `Gains` (IS, sign so that positive = gain), `EquityEarnings` (IS), `DistributionsFromAffiliates` (CF), `DistributionsToNCI` (CF), `NetPPE`, `EquityInvestments`, `Receivables` (`AccountsReceivableNetCurrent`, `ReceivablesNetCurrent`), `Inventory` (`InventoryNet`), `Payables` (`AccountsPayableCurrent`, `AccountsPayableTradeCurrent`), `TotalAssets`, `LiabilitiesAndEquity` (`LiabilitiesAndStockholdersEquity`, `LiabilitiesAndPartnersCapital`), `TotalEquityInclNCI` (`StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest`, `PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest`), `NCI` (`MinorityInterest`), `PretaxIncome` (`IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest`, `IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments`), `DeferredTax` (CF, `DeferredIncomeTaxExpenseBenefit`), `InvestingCF`, `FinancingCF`, `NetChangeInCash` (`CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect`, `CashAndCashEquivalentsPeriodIncreaseDecrease`), `CFBeginningCash`/`CFEndingCash` (same tags; beginning = the row whose label matches `beginning`, ending = `end`), `CFNetIncome` (CF, `ProfitLoss`, `NetIncomeLoss`), `DebtCurrent`, `LongTermDebt`.
  - **Canonical concept names (exact strings, used by Task 9):** `Revenue`, `TotalCostsAndExpenses`, `OperatingIncome`, `DA`, `Impairments`, `Gains`, `EquityEarnings`, `InterestExpense`, `PretaxIncome`, `IncomeTax`, `NetIncome`, `NIToNCI`, `DilutedUnits`, `UnitsOutstanding`, `OperatingCashFlow`, `CFNetIncome`, `DeferredTax`, `DistributionsFromAffiliates`, `Capex`, `InvestingCF`, `Distributions`, `DistributionsToNCI`, `FinancingCF`, `NetChangeInCash`, `CFBeginningCash`, `CFEndingCash`, `Cash`, `Receivables`, `Inventory`, `NetPPE`, `EquityInvestments`, `TotalAssets`, `Payables`, `DebtCurrent`, `LongTermDebt`, `TotalDebt`, `LiabilitiesAndEquity`, `EquityCommon`, `NCI`, `TotalEquityInclNCI`.
  - `mapping.py`: `resolve(statements: dict[str, MergedStatement]) -> dict[str, dict[str, ConceptValue]]` → `concept -> fy -> ConceptValue`. Per concept per year: first tag in `tags` order whose row (in the concept's statement) has a non-`None` value that year; else label regex over non-abstract rows of that statement; else `UNMAPPED`. `sum_tags`: sum of all found, status `DERIVED`. `TotalDebt` = `LongTermDebtNoncurrent + LongTermDebtCurrent + DebtCurrent + CommercialPaper + ShortTermBorrowings` if any present, else `LongTermDebt`.
  - `coverage_warnings(mapping) -> list[str]` → one line per concept in `REQUIRED = ("Revenue","NetIncome","TotalAssets","LiabilitiesAndEquity","OperatingCashFlow","Cash","TotalDebt","EquityCommon","DilutedUnits")` that is `UNMAPPED` in the latest year.

- [ ] **Step 1: Write failing tests** (hand-built `MergedStatement`s)

```python
def test_tag_order_preferred_over_label(): ...
def test_label_fallback_for_custom_tag():            # EPD D&A via epd_ tag + label "Depreciation and accretion"
def test_per_year_resolution_can_change_tag():       # FY2023 standard tag, FY2024 custom tag -> both OK, tags differ
def test_total_debt_sums_components(): ...
def test_unmapped_reported_and_required_flagged(): ...
def test_cf_beginning_vs_ending_cash_by_label(): ...
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit** — `git commit -am "feat: concept mapping table and resolver"`

---

### Task 7: Checks and companyfacts cross-check

**Files:**
- Create: `sec_model/checks.py`, `sec_model/companyfacts.py`
- Test: `tests/test_checks.py`

**Interfaces:**
- Produces:
  - `companyfacts.py`: `load_companyfacts(client, cik) -> dict`; `unit_hint_from_facts(facts) -> Callable[[str], str | None]` (maps `us-gaap_X` → `"money"` if units key is a 3-letter currency, `"shares"` if `shares`, `"perunit"` if contains `/`, else `None`); `fy_values(facts, tag: str, currency: str) -> dict[str, float]` → `fy_label -> value in millions` using entries with `fp == "FY"`, form in 10-K family, 12-month duration (`end - start` between 350 and 380 days, or instant for BS tags), newest `filed` wins; `dei_shares_outstanding(facts) -> dict[str, float]` (millions, by FY of `fy`).
  - `checks.py`: `run_checks(statements, mapping, facts_values: dict[str, dict[str, float]]) -> list[CheckResult]` implementing spec §5.5 exactly: tolerance `max(1.0, 0.005 * max(|lhs|, |rhs|))`; API cross-check `WARN` if outside tolerance, `FAIL` if relative diff > 5%; `SKIP` with a note when an input is unmapped. Subtotal check: for each non-abstract row whose label starts with `Total` (case-insensitive), children = contiguous preceding non-abstract rows with `indent > total.indent` back to the previous abstract/row at the same or lower indent; `WARN` when their sum differs; skip if any child is `None`.
  - `summarize(results) -> dict[str, int]` → counts by status.

- [ ] **Step 1: Write failing tests** (synthetic statements: a balanced one, one off by 2, one with restricted cash difference)

```python
def test_balance_sheet_pass_within_tolerance(): ...     # off by 0.4 on 10,000 -> PASS
def test_balance_sheet_fail(): ...                       # off by 100 -> FAIL
def test_ni_tie_is_vs_cf(): ...
def test_cash_roll_forward(): ...
def test_cf_vs_bs_cash_is_warn_only(): ...
def test_api_crosscheck_fail_above_5pct(): ...
def test_unmapped_input_gives_skip(): ...
def test_subtotal_children_sum_warn(): ...
def test_fy_values_picks_12_month_fy_entries_newest_filed(): ...   # synthetic companyfacts dict
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit** — `git commit -am "feat: tie-out checks and companyfacts cross-check"`

---

### Task 8: Formula AST (one definition, two renderers)

**Files:**
- Create: `sec_model/formula.py`
- Test: `tests/test_formula.py`

**Interfaces:**
- Produces:
  - Expression nodes (all subclasses of `Expr` with `+ - * /` and unary `-` overloaded, returning `BinOp`/`Neg`): `Const(value: float)`, `Line(key: str, lag: int = 0)` (another Model line in the same column, or `lag=1` → prior column), `Driver(key: str)` (per-year driver cell in the same column on the Drivers sheet), `Scalar(key: str)` (single named cell, e.g. `JV_TOGGLE`, `MIN_CASH`), `Block(key: str)` (per-column SUM over a named block of rows on another sheet, e.g. asset revenue block), `Fn(name: str, *args: Expr)` for `MAX`, `MIN`, `IF`, `AVERAGE`, `ABS`, `IFERROR`; `Hist(key: str, n: int)` → the `AVERAGE` of the line's last `n` history columns (renders to an absolute range).
  - `to_excel(expr, ctx: RenderContext) -> str` where `RenderContext(col: int, line_cell: Callable[[str, int], str], driver_cell: Callable[[str, int], str], scalar_cell: Callable[[str], str], block_range: Callable[[str, int], str], hist_range: Callable[[str, int], str])` returns e.g. `"=(D12-D13)+Drivers!E7*D20"`. `Line` with `lag` into a column before the first history column renders `0`.
  - `evaluate(expr, env: EvalContext) -> float` where `EvalContext(col: int, line: Callable[[str, int], float], driver, scalar, block, hist)`. `None` values evaluate as `0.0`. Division by zero → `0.0` (and `to_excel` wraps `/` in `IFERROR(...,0)`).

- [ ] **Step 1: Write failing tests**

```python
def test_arithmetic_renders_and_evaluates():
    e = (Line("REV") - Line("COST")) * Driver("MARGIN")
    assert to_excel(e, ctx) == "=(D5-D6)*Drivers!D9"
    assert evaluate(e, env) == pytest.approx((100-60)*0.5)

def test_lag_before_first_column_is_zero(): ...
def test_div_by_zero_is_zero_and_iferror_in_excel(): ...
def test_fn_max_if_and_block(): ...
def test_hist_average_range(): ...   # "AVERAGE($D$5:$F$5)"
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit** — `git commit -am "feat: formula AST with Excel and Python renderers"`

---

### Task 9: Model specification (history derivations, drivers, forecast) + Python evaluation

**Files:**
- Create: `sec_model/model_spec.py`, `sec_model/forecast.py`
- Test: `tests/test_forecast.py`

**Interfaces:**
- Produces:
  - `model_spec.py`:
    - `@dataclass LineDef: key; label; section ("IS"|"CF"|"BS"|"CHECK"); hist: str | Expr | None` (a concept name → history cells pull from the Mapping sheet; an `Expr` → derived from other lines in history columns; `None` → blank in history) `; fcst: Expr | None; fmt: str ("money"|"pct"|"units"|"perunit"); bold: bool = False; is_input_row: bool = False`
    - `@dataclass DriverDef: key; label; default: Expr` (an expression over history lines/Hist averages that fills the driver's first forecast column; later columns default to `=prev column`) `; fmt; note: str`
    - `@dataclass ScalarDef: key; label; default: float | Expr; fmt; note`
    - `MODEL_LINES: list[LineDef]`, `DRIVERS: list[DriverDef]`, `SCALARS: list[ScalarDef]`, built by `build_spec(da_basis: str) -> ModelSpec` where `da_basis in ("ppe", "revenue")`.
    - The lines, in order, with their definitions (sign convention: expenses positive on IS lines named as expenses; CF lines carry their cash sign):

      **IS**: `REV` (hist concept `Revenue`; fcst `Block("ASSET_REV")`) · `OPCOST` (hist `TotalCostsAndExpenses − DA − IMPAIR − (−GAINS)`, falling back to `REV − OperatingIncome − DA − IMPAIR + GAINS` when `TotalCostsAndExpenses` unmapped; fcst `Block("ASSET_COST")`) · `ASSET_EBITDA = REV − OPCOST` (bold) · `EQ_EARN` (hist `EquityEarnings`; fcst `Line("EQ_EARN",1)`) · `EBITDA = ASSET_EBITDA + EQ_EARN*Scalar("JV_TOGGLE")` (bold) · `DA` (hist `DA`; fcst `Driver("DA_PCT") * Line("NET_PPE",1)` or `* Line("REV")` by `da_basis`) · `IMPAIR` (hist `Impairments`; fcst 0) · `GAINS` (hist `Gains`; fcst 0) · `EBIT = EBITDA + EQ_EARN*(1−JV_TOGGLE) − DA − IMPAIR + GAINS` (bold) · `INTEREST` (hist `InterestExpense`; fcst `Driver("INT_RATE")*(Line("DEBT",1)+Line("REVOLVER",1))`) · `OTHER_INC` (hist `PretaxIncome − EBIT + INTEREST` **derived**; fcst `Line("OTHER_INC",1)`) · `PRETAX = EBIT − INTEREST + OTHER_INC` · `TAX` (hist `IncomeTax`; fcst `Driver("TAX_RATE")*PRETAX`) · `DISC_OPS` (hist `NetIncome − (PRETAX − TAX)` derived; fcst 0) · `NI = PRETAX − TAX + DISC_OPS` (bold) · `NI_NCI` (hist `NIToNCI`; fcst lag) · `NI_COMMON = NI − NI_NCI` (bold) · `UNITS` (hist `DilutedUnits`; fcst `Driver("UNITS")`) · `EPS = NI_COMMON / UNITS` (perunit).

      **CF**: `CF_NI = NI` · `CF_DA = DA` · `CF_IMPAIR = IMPAIR` · `CF_GAINS = −GAINS` · `CF_EQ = −EQ_EARN` · `CF_DIST_AFF` (hist `DistributionsFromAffiliates`; fcst lag) · `CF_DWC = −((AR − AR₋₁) + (INV − INV₋₁) − (AP − AP₋₁))` for both hist and fcst · `CF_OTHER_OP` (hist `OperatingCashFlow − sum of the above` derived; fcst `Driver("OTHER_OP")`) · `OCF` (bold, sum) · `CAPEX` (hist `−|Capex|`; fcst `−Driver("CAPEX_PCT")*REV`) · `CF_OTHER_INV` (hist `InvestingCF − CAPEX` derived; fcst 0) · `ICF` (bold) · `DIST` (hist `−|Distributions|`; fcst `−Driver("DPU")*UNITS`) · `DIST_NCI` (hist `−|DistributionsToNCI|`; fcst `−NI_NCI`) · `DEBT_NET` (hist `DEBT − DEBT₋₁` derived; fcst `Driver("DEBT_ISSUE")`) · `REVOLVER_NET` (hist 0; fcst `MAX(Scalar("MIN_CASH") − CASH_PRE, −REVOLVER₋₁)` where `CASH_PRE = CASH₋₁ + OCF + ICF + DIST + DIST_NCI + DEBT_NET`) · `CF_OTHER_FIN` (hist `FinancingCF − DIST − DIST_NCI − DEBT_NET` derived; fcst 0) · `FCF_FIN` (bold) · `FX_OTHER` (hist `NetChangeInCash − OCF − ICF − FCF_FIN` derived; fcst 0) · `NET_CHANGE = OCF + ICF + FCF_FIN + FX_OTHER` (bold).

      **BS**: `CASH` (hist `Cash`; fcst `CASH₋₁ + NET_CHANGE`) · `AR` (hist `Receivables`; fcst `Driver("AR_DAYS")/365*REV`) · `INV` (hist; fcst `Driver("INV_DAYS")/365*OPCOST`) · `NET_PPE` (hist; fcst `NET_PPE₋₁ − CAPEX − DA`) · `EQ_INV` (hist; fcst lag) · `OTHER_ASSETS` (hist `TotalAssets − CASH − AR − INV − NET_PPE − EQ_INV` derived; fcst lag) · `TOTAL_ASSETS` (bold, sum) · `AP` (hist `Payables`; fcst `Driver("AP_DAYS")/365*OPCOST`) · `DEBT` (hist `TotalDebt`; fcst `DEBT₋₁ + DEBT_NET`) · `REVOLVER` (hist 0; fcst `REVOLVER₋₁ + REVOLVER_NET`) · `OTHER_LIAB` (hist `LiabilitiesAndEquity − TotalEquityInclNCI − AP − DEBT` derived, with `TotalEquityInclNCI` falling back to `EquityCommon + NCI`; fcst lag) · `EQUITY` (hist `EquityCommon`; fcst `EQUITY₋₁ + NI_COMMON + DIST`) · `NCI_BS` (hist `NCI`; fcst `NCI_BS₋₁ + NI_NCI + DIST_NCI`) · `TOTAL_LE` (bold, sum) · **CHECK** `BAL_CHECK = TOTAL_ASSETS − TOTAL_LE`.
      **RATIO** (history and forecast columns, shown under the IS as "Ratios & driver history"; these are what `Hist(...)` averages): `REV_GROWTH_HIST = REV/REV₋₁ − 1` · `EBITDA_MARGIN = EBITDA/REV` · `COST_PCT_HIST = OPCOST/REV` · `DA_PCT_HIST = DA/NET_PPE₋₁` (or `DA/REV` by `da_basis`) · `CAPEX_PCT_HIST = −CAPEX/REV` · `INT_RATE_HIST = INTEREST/AVERAGE(DEBT, DEBT₋₁)` · `TAX_RATE_HIST = MAX(0, MIN(0.35, (TAX − DEFTAX)/PRETAX))` where `DEFTAX` is a hidden history line (concept `DeferredTax`, 0 when unmapped) · `DPU_HIST = −DIST/UNITS` · `DPU_GROWTH_HIST = DPU_HIST/DPU_HIST₋₁ − 1` · `AR_DAYS_HIST = AR/REV*365` · `INV_DAYS_HIST = INV/OPCOST*365` · `AP_DAYS_HIST = AP/OPCOST*365` · `NET_DEBT_EBITDA = (DEBT + REVOLVER − CASH)/EBITDA` · `DIST_COVERAGE = (OCF + CAPEX)/−DIST`.
    - Drivers: `REV_GROWTH` (default `Hist("REV_GROWTH_HIST",3)` capped ±15% via `MAX(MIN(..,0.15),-0.15)`; used only by the placeholder asset), `COST_PCT` (last year `OPCOST/REV`), `DA_PCT` (`DA/NET_PPE₋₁` or `DA/REV` 3-yr avg), `CAPEX_PCT` (`|CAPEX|/REV` 3-yr avg), `INT_RATE` (last `INTEREST/AVERAGE(DEBT, DEBT₋₁)`), `TAX_RATE` (`MAX(0,MIN(0.35,(TAX − DeferredTax)/PRETAX))` 3-yr avg; `DeferredTax` 0 when unmapped), `DPU` (last `|DIST|/UNITS` grown by `DPU_GROWTH`), `DPU_GROWTH` (3-yr CAGR of DPU capped ±15%), `UNITS` (last), `AR_DAYS`/`INV_DAYS`/`AP_DAYS` (last year), `OTHER_OP` (0), `DEBT_ISSUE` (0).
    - Scalars: `JV_TOGGLE` (1; note "1 = include JV equity earnings in EBITDA, 0 = exclude"), `WC_HOLD` (0; note "1 = hold working capital flat (no change in forecast)"; `CF_DWC` forecast is multiplied by `(1 − WC_HOLD)` and AR/INV/AP forecast become `lag` when `WC_HOLD = 1`, via `IF`), `MIN_CASH` (last `CASH`).
  - `forecast.py`: `compute(spec: ModelSpec, history: dict[str, dict[str, float | None]], fy_labels: list[str], fcst_labels: list[str], driver_overrides: dict | None = None) -> dict[str, dict[str, float]]` → every line × every column (history columns use `hist`, forecast columns use `fcst`) evaluated with Task 8's `evaluate`; `asset_block` defaults: one asset with `REV × (1+REV_GROWTH)^n` and `COST_PCT × REV`. Also `driver_defaults(spec, history, fy_labels) -> dict[str, float]`.

- [ ] **Step 1: Write failing tests** (a hand-built 3-year history for a toy company with round numbers; a second history shaped like EPD: no `TotalCostsAndExpenses`? — no, EPD has it; shape it like OKE: no `Liabilities`, no `TotalCurrentLiabilities`)

```python
def test_history_balance_check_zero_with_derived_other_liab(): ...   # Review Focus 4
def test_forecast_balances_every_year(): assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)
def test_revolver_draws_shortfall_to_min_cash():                      # Review Focus 5
    # set DIST huge so cash would go below MIN_CASH; assert CASH == MIN_CASH and REVOLVER == shortfall
def test_revolver_repays_when_cash_exceeds_min(): ...
def test_jv_toggle_moves_ebitda_not_net_income(): ...
def test_da_basis_revenue_when_ppe_missing(): ...
def test_nci_flat_by_default(): ...
def test_tax_rate_uses_deferred_tax_when_present(): ...
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit** — `git commit -am "feat: model spec and python forecast evaluator"`

---

### Task 10: Workbook — styles, Cover, Checks, As-Reported, Mapping, Sources

**Files:**
- Create: `sec_model/workbook/__init__.py`, `sec_model/workbook/styles.py`, `sec_model/workbook/asreported.py`, `sec_model/workbook/meta.py` (Cover/Checks/Sources), `sec_model/workbook/mapping_sheet.py`
- Test: `tests/test_workbook_asreported.py`

**Interfaces:**
- Consumes: `MergedStatement`, mapping dict, `CheckResult`s
- Produces:
  - `styles.py`: constants from Global Constraints (`NAVY="1F3864"`, fills, fonts, `FMT_MONEY='#,##0.0;(#,##0.0);"-"'`, `FMT_PCT='0.0%;(0.0%);"-"'`, `FMT_UNITS`, `FMT_PERUNIT='#,##0.00;(#,##0.00);"-"'`, `FMT_X='0.0"x"'`), helpers `title_bar(ws, row, text, last_col)`, `header_row(ws, row, labels, first_col, is_forecast: list[bool])`, `set_widths(ws, {...})`.
  - `asreported.py`: `write_as_reported(wb, stmt: MergedStatement, sheet_name: str) -> dict[str, int]` → row number per `MergedRow.key`. Layout: col A = hidden row key (`ws.column_dimensions["A"].hidden = True`), B = label (indented by `indent`), C = unit note (`$MM`/`MM units`/`per unit`), D.. = FY columns oldest→newest. Abstract rows bold with light-gray fill; rows whose label starts with `Total` bold with a thin top border; `still_reported=False` rows italic gray with `" (no longer reported)"` appended to the label. Freeze panes `D5`. Sheet names: `IS (As Reported)`, `BS (As Reported)`, `CF (As Reported)`.
  - `mapping_sheet.py`: `write_mapping(wb, mapping, stmt_sheets: dict[str, str], row_index: dict[str, dict[str, int]], fy_labels, concept_stype: dict[str, str]) -> dict[str, int]` → row per concept. Columns: B concept, C status (latest year), D.. per-FY **value** cells as formulas `=IFERROR(INDEX('<sheet>'!<col>:<col>, MATCH(<keycell>, '<sheet>'!$A:$A, 0)), "")`, and to the right a per-FY **key** block (plain text, yellow) holding the row key used that year (blank when unmapped), then tag and label (latest year). This is what the analyst edits to repoint.
  - `meta.py`: `write_cover(wb, info: CoverInfo)` (company, ticker, CIK, FYE, currency, years, filing-per-year table, check counts, sheet index with hyperlinks, pulled date, data-source note); `write_checks(wb, results)`; `write_sources(wb, statements)`.

- [ ] **Step 1: Write failing tests** (build a workbook from a hand-made `MergedStatement`, reopen with openpyxl, assert)

```python
def test_as_reported_layout_and_hidden_key_column(): ...
def test_dropped_row_label_and_style(): ...
def test_mapping_value_formula_points_at_key(): ...   # formula string contains "MATCH(" and the key cell ref
def test_cover_lists_filing_per_year_and_check_counts(): ...
def test_workbook_reopens_without_error(): ...
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit**

---

### Task 11: Workbook — Model, Asset Inputs, Drivers sheets

**Files:**
- Create: `sec_model/workbook/model_sheet.py`
- Test: `tests/test_workbook_model.py`

**Interfaces:**
- Consumes: `ModelSpec` (Task 9), Task 8 renderers, Mapping row index (Task 10)
- Produces: `write_model(wb, spec, mapping_rows: dict[str, int], fy_labels, fcst_labels, driver_defaults: dict[str, float], scalar_defaults: dict[str, float], currency: str) -> ModelLayout` where `ModelLayout(line_row: dict[str,int], driver_row, scalar_cell: dict[str,str], first_hist_col: int, first_fcst_col: int, sheet="Model")`.
  - **Drivers** sheet: one row per `DriverDef` with forecast-year columns aligned to the Model's forecast columns (same column letters); first forecast cell = the Python-computed default **value** (blue/yellow input), later cells `=<prev col>` (so editing year 1 flows right; an analyst can overwrite any year). Scalars block below (`JV_TOGGLE`, `MIN_CASH`) as named cells (`wb.defined_names`). Each driver has a `note` column explaining the default.
  - **Asset Inputs** sheet: 25 asset rows × (name, segment, revenue by forecast year, operating cost by forecast year); row 1 pre-filled `Consolidated` with formulas `=Model!<REV last hist>*(1+Drivers!<REV_GROWTH>)^n` and `=<rev cell>*Drivers!<COST_PCT>`; rows 2–24 blank yellow; row 25 `Other / unallocated` blank. Totals rows below each block (`SUM` of the 25). Reconciliation rows: last reported revenue/opcost (links to Model) and `Asset EBITDA`. Defined names `ASSET_REV_<FY>`/`ASSET_COST_<FY>` are not needed; `Block("ASSET_REV")` renders to the total cell for that column.
  - **Model** sheet: sections IS / CF / BS / CHECK with the lines in `MODEL_LINES` order; history cells per `LineDef.hist` (concept → `=Mapping!<cell>`; `Expr` → rendered; `None` → blank), forecast cells rendered from `fcst`; `BAL_CHECK` row red conditional format when `ABS > 0.5`; a ratios block under the IS (`EBITDA margin`, `Revenue growth`, `Net debt / EBITDA`, `Distribution coverage = (OCF + CAPEX) / −DIST`) as formulas; history columns from the Mapping sheet, forecast columns `EBF3FB`.

- [ ] **Step 1: Write failing tests**

```python
def test_model_history_cell_links_to_mapping(): ...
def test_forecast_cell_formula_matches_renderer():   # pick EBITDA: cell formula == to_excel(spec EBITDA fcst, ctx)
def test_driver_first_col_value_and_chain(): ...
def test_asset_block_total_feeds_revenue(): ...      # Model REV forecast cell formula references 'Asset Inputs'!<total cell>
def test_balance_check_has_conditional_format(): ...
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit**

---

### Task 12: Workbook — WACC and DCF

**Files:**
- Create: `sec_model/workbook/valuation.py`
- Test: `tests/test_workbook_valuation.py`

**Interfaces:**
- Consumes: `ModelLayout`, `MarketInputs`
- Produces: `write_wacc(wb, layout, market, currency) -> WaccLayout(wacc_cell, tax_cell, ...)` and `write_dcf(wb, layout, wacc, market, fcst_labels, currency, basic_units: float | None)` (`basic_units` = latest `dei_shares_outstanding` in millions from Task 7, falling back to Model `UNITS`). When `currency != "USD"` both sheets show, under the title, `f"All figures in {currency} MM. Enter share price in {currency}."` in red italic.
  - **WACC** (mirror template rows): Kd input (default = Drivers `INT_RATE` link), tax rate (= Drivers `TAX_RATE` link), Kd after tax; Rf/β/MRP inputs from `market`; Ke; capital structure: equity = `price × diluted units` (price cell yellow, blank-safe), net debt = `Model` last-year `DEBT − CASH`; target override column; weights; WACC (green). Beta selector cell `Observed`/`Industry` with a 6-row peer table (ticker, β, price, shares, net debt, tax) → unlevered β per row, average, relevered; `IF(selector="Observed", observed, relevered)`. Sensitivity grid WACC vs (β ± 0.1, ±0.2) × (MRP ± 1%, ±2%) as formulas `=Kd_at*Wd + (Rf + β_i*MRP_j)*We`.
  - **DCF** (mirror template): header (company, ticker, FYE, valuation date `=TODAY()`, stub `=YEARFRAC(TODAY(), first forecast FYE)`, mid-year toggle 0/1, WACC link, price link); UFCF block linking Model forecast lines: `EBITDA`, `EBIT`, tax on EBIT `=EBIT*TAX_RATE`, NOPAT, `DA`, ΔNWC (`CF_DWC`), capex, UFCF; discount dates per template (`IF(mid=0, FYE, AVERAGE(prev FYE, FYE))`); PV factors `=(1+WACC)^-((date − valuation)/365)`; stub-adjusted year-1 UFCF. Terminal value both ways: perpetuity (`g` input default 0.025) and exit multiple (input default `=IFERROR((price×units + net debt)/last-year EBITDA, 10)`); PV of TV; EV; TV %; implied multiple / implied g. Net-debt bridge rows (debt, revolver, preferred 0, NCI, less cash, less equity investments) from Model with override cells. Dilution block (basic units from cover count; RSU/options/convertibles inputs 0) → diluted units → equity value → value per unit under both methods; `% over/(under)` = `IFERROR(price/value − 1, "enter price")`. Multiples block (Current / Year-1 / Terminal: EV/Revenue, EV/EBITDA, EV/EBIT).
  - **Sensitivity grids as formulas**: a helper block computes `PV of UFCF at WACC w` and `TV at (w, g)` / `TV at (w, m)` in closed form per cell: `SUMPRODUCT(UFCF_range, (1+w)^-(t_range))` where `t_range` is a row of year fractions; value per unit = `(PV + TV_pv − net debt)/units`. Grids: value per unit for WACC (−1%, −0.5%, 0, +0.5%, +1%) × g (−1%…+1%) and WACC × exit multiple (−1.0x…+1.0x). Football-field table: low/high per method, plus 52-week low/high inputs; a stacked `BarChart` (`type="bar"`, `grouping="stacked"`, first series no fill) drawn from it — **cut candidate**; if cut, leave the table.

- [ ] **Step 1: Write failing tests**

```python
def test_wacc_formula_references_weights_and_costs(): ...
def test_dcf_ufcf_rows_link_to_model(): ...           # EBITDA cell formula == "=Model!<col><row>"
def test_sensitivity_grid_cells_are_formulas_not_data_tables(): ...   # no cell has a DataTableFormula; grid center formula contains "SUMPRODUCT("
def test_exit_multiple_default_falls_back_without_price(): ...
def test_no_ref_errors_anywhere(): ...                # scan all cells for "#REF!"
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit**

---

### Task 13: Pipeline and `build_workbook`

**Files:**
- Create: `sec_model/pipeline.py`, `sec_model/workbook/build.py`
- Test: `tests/test_pipeline.py` (offline, using fixtures through a `FakeClient` that serves fixture files for the fixture ticker's URLs)

**Interfaces:**
- Produces:
  - `@dataclass BuildResult: company: str; ticker: str; cik: int; currency: str; fy_labels: list[str]; fcst_labels: list[str]; statements: dict[str, MergedStatement]; mapping; checks: list[CheckResult]; coverage: list[str]; forecast: dict; xlsx: bytes; log: list[str]`
  - `build(ticker: str, user_agent: str, n_years: int = 10, n_fcst: int = 5, market: MarketInputs = MarketInputs(), progress: Callable[[str], None] = print, cache_dir=Path(".cache")) -> BuildResult`. Steps in order: resolve → filings → per FY `choose_source_filing` → parse the three R-files of each chosen filing with `unit_hint` from companyfacts → merge → mapping → checks → `da_basis = "ppe" if NET_PPE mapped in latest year else "revenue"` → `build_spec` → `driver_defaults` → `compute` (for the regression assertion and the app preview) → `build_workbook` → result. Each step appends to `log` and calls `progress`.
  - `build_workbook(result-ish args) -> bytes` in `workbook/build.py`, sheet order: Cover, Checks, Model, Asset Inputs, Drivers, WACC, DCF, Mapping, IS/BS/CF (As Reported), Sources.
  - Fiscal years with no filing that has all three statements are dropped with a log line; if fewer than 2 remain → `UserFacingError("Fewer than two fiscal years with full XBRL statements found.")`.

- [ ] **Step 1: Write failing tests**

```python
def test_build_offline_fixture_ticker_produces_workbook(): ...   # len(r.xlsx) > 0; sheetnames in order; r.checks non-empty
def test_build_unknown_ticker_user_facing(): ...
def test_forecast_balances_in_result(): ...
```

- [ ] **Step 2–4: FAIL → implement → PASS** → **Step 5: Commit**

---

### Task 14: Streamlit app, shim, README

**Files:**
- Create: `app.py`
- Modify: `streamlit_app.py` → two lines: `from app import main` / `main()` (guarded by `if __name__ == "__main__"` is not needed for Streamlit; keep it a plain import-and-run), `README.md`, `.devcontainer` (leave), `requirements.txt` (unchanged content, confirm)
- Test: manual (Streamlit has no unit harness here)

- [ ] **Step 1: Implement `app.py`** per spec §4: sidebar (ticker, email, years 2–15 default 10, forecast 3–7 default 5, "Market inputs" expander: price (blank allowed), beta, Rf, MRP; Build button); `st.status` progress fed by `progress`; checks table with colored status (`st.dataframe` with a style map); preview tabs from `BuildResult.statements` + Mapping; download button `f"{ticker}_3Stmt_{date}.xlsx"`. `UserFacingError` → `st.error(str(e))`; any other exception → `st.error("Build failed; see server log.")` + `logging.exception`. Wrap `build` in `st.cache_data(ttl=3600)` keyed on all inputs.
- [ ] **Step 2: Run** `streamlit run app.py`, build OKE with your email, download, open in Excel: no repair prompt, no `#REF!`, change `Asset Inputs` row 1 revenue → `Model` EBITDA, `CASH`, `REVOLVER`, and both DCF values change. Note the result in the commit message.
- [ ] **Step 3: Update README** (what it does, run locally, inputs, sheets list, checks list, limitations incl. 40-F and CAD).
- [ ] **Step 4: Commit**

---

### Task 15: Regression script and acceptance run

**Files:**
- Create: `scripts/regress.py`
- Test: the script itself

- [ ] **Step 1: Implement** `python scripts/regress.py --ua "Name email" OKE EPD ET KMI WMB MPLX TRGP KNTK ENB AAPL XOM JPM TRP`: for each ticker calls `pipeline.build`, writes `out/<TICKER>.xlsx`, and prints one line: ticker, years found, statements found per year, check counts, coverage warnings, `max |BAL_CHECK|` over forecast years, elapsed. Asserts (collecting failures, exit code 1 at the end): no exception except `UserFacingError` for `TRP` (which must mention `40-F`); ≥ 5 years; all three statements every kept year; zero `FAIL` checks; `max |BAL_CHECK| < 1e-6`; `openpyxl.load_workbook(out)` succeeds; no `#REF!` in any cell.
- [ ] **Step 2: Run it**; for each failing ticker apply the systematic-debugging skill, fix in the owning module with a test, re-run until the full list passes. Record the final output table in `docs/superpowers/plans/2026-10-22-acceptance-run.md`.
- [ ] **Step 3: Manual Excel check** on OKE and EPD per Task 14 Step 2; record in the same file.
- [ ] **Step 4: Commit** — `git commit -am "test: regression script and acceptance run"`
