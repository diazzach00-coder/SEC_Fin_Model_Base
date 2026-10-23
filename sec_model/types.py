from dataclasses import dataclass, field
from datetime import date


@dataclass(frozen=True)
class Filing:
    cik: int
    accession: str          # "0001039684-26-000012"
    form: str               # "10-K" | "10-K/A" | "10-KT"
    filing_date: date
    report_date: date       # fiscal year end
    primary_doc: str

    @property
    def fy_label(self) -> str:
        return "FY" + str(self.report_date.year)

    @property
    def folder_url(self) -> str:
        return f"https://www.sec.gov/Archives/edgar/data/{self.cik}/{self.accession.replace('-', '')}/"


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
    values: list            # list[float | None], one per Column


@dataclass
class ParsedStatement:
    stype: str              # "IS" | "BS" | "CF"
    title: str
    currency: str           # "USD" | "CAD" ...
    money_scale: float      # multiplier to millions: 1 for "$ in Millions", 0.001 for thousands
    share_scale: float
    columns: list           # list[Column]
    rows: list              # list[Row]
    source_url: str


@dataclass
class SourceRec:
    accession: str
    form: str
    filing_date: date
    r_url: str
    scale_text: str


@dataclass
class MergedRow:
    key: str
    tag: str
    label: str
    indent: int
    is_abstract: bool
    unit: str
    values: dict            # dict[str, float | None]  fy_label -> value
    still_reported: bool


@dataclass
class MergedStatement:
    stype: str
    currency: str
    fy_labels: list         # oldest -> newest
    fy_ends: dict           # fy_label -> date
    rows: list              # list[MergedRow]
    sources: dict           # fy_label -> SourceRec


@dataclass
class ConceptValue:
    concept: str
    fy: str
    value: float | None
    row_key: str | None
    tag: str | None
    label: str | None
    status: str             # "OK" | "UNMAPPED" | "DERIVED"
    # Per-cell sign applied to the As-Reported value (on top of the concept's sign): -1 for an
    # income-positive tag mapped to an expense concept, or a negated-label row whose
    # presentation sign is opposite to the XBRL fact. The Mapping sheet applies it per year.
    sign: int = 1


@dataclass
class CheckResult:
    name: str
    fy: str
    status: str             # "PASS" | "WARN" | "FAIL" | "SKIP"
    lhs: float | None
    rhs: float | None
    tolerance: float
    note: str


@dataclass
class MarketInputs:
    price: float | None = None
    beta: float = 1.0
    risk_free: float = 0.0425
    mrp: float = 0.055


class UserFacingError(Exception):
    """Message is safe to show to the end user verbatim."""
