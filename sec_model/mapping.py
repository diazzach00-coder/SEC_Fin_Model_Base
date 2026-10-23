"""Resolve each model concept to one as-reported row per fiscal year."""
import re
from collections.abc import Callable
from dataclasses import replace

from .concepts import CONCEPTS, Concept
from .types import ConceptValue, MergedRow, MergedStatement

REQUIRED = ("Revenue", "NetIncome", "TotalAssets", "LiabilitiesAndEquity",
            "OperatingCashFlow", "Cash", "TotalDebt", "EquityCommon", "DilutedUnits")

# These share their tags, so a tag hit must also satisfy the label regex.
_LABEL_FILTERED = frozenset({"CFBeginningCash", "CFEndingCash"})


def _is_undimensioned(r: MergedRow) -> bool:
    """True for the consolidated row: key is the tag, or the parser's ``tag|label``
    (``|n``) key for a repeated undimensioned tag. Dimensional keys carry a member."""
    if r.key == r.tag:
        return True
    base = f"{r.tag}|{r.label}"
    return r.key == base or re.fullmatch(re.escape(base) + r"\|\d+", r.key) is not None


def _usable(r: MergedRow, fy: str) -> bool:
    return not r.is_abstract and r.values.get(fy) is not None


def _unmapped(concept: str, fy: str) -> ConceptValue:
    return ConceptValue(concept, fy, None, None, None, None, "UNMAPPED")


def _ok(c: Concept, fy: str, r: MergedRow) -> ConceptValue:
    ts = c.tag_sign(r.tag)
    return ConceptValue(c.name, fy, c.sign * ts * r.values[fy], r.key, r.tag, r.label, "OK",
                        sign=ts)


def _resolve_sum(c: Concept, rows: list[MergedRow], fy: str) -> ConceptValue | None:
    used: list[MergedRow] = []
    for g in c.sum_groups:
        if any(r.tag.removeprefix("us-gaap_") in g.skip_if for r in used):
            continue
        for t in g.tags:
            hit = next((r for r in rows if r.key == f"us-gaap_{t}" and _usable(r, fy)), None)
            if hit is not None:
                used.append(hit)
                if g.mode == "first":
                    break
    if not used:
        return None
    if len(used) == 1:                    # one row: keep it linked like a tag match
        return _ok(c, fy, used[0])
    return ConceptValue(
        c.name, fy, c.sign * sum(r.values[fy] for r in used), None,
        "SUM(" + "+".join(r.tag for r in used) + ")",
        " + ".join(r.label for r in used), "DERIVED")


def _resolve_by_tag(c: Concept, stmt: MergedStatement, fy: str) -> ConceptValue | None:
    rows = stmt.rows
    label_re = re.compile(c.label_re, re.IGNORECASE) if c.label_re else None

    if c.sum_groups:
        derived = _resolve_sum(c, rows, fy)
        if derived is not None:
            return derived

    filtered = c.name in _LABEL_FILTERED
    for t in c.tags:
        full = f"us-gaap_{t}"
        for r in rows:
            if r.tag != full or not _usable(r, fy):
                continue
            if filtered:
                if _is_undimensioned(r) and label_re.search(r.label):
                    return _ok(c, fy, r)
            elif r.key == full:
                return _ok(c, fy, r)

    if filtered:
        # Labels like "..., January 1" / "..., December 31" carry no keyword: the
        # first undimensioned cash row is the opening balance, the last the closing.
        for t in c.tags:
            same = [r for r in rows if r.tag == f"us-gaap_{t}" and _is_undimensioned(r)
                    and not r.is_abstract]
            if len(same) >= 2:
                r = same[0] if c.name == "CFBeginningCash" else same[-1]
                if r.values.get(fy) is not None:
                    return _ok(c, fy, r)
    return None


def _add_on(c: Concept, stmt: MergedStatement, fy: str, cv: ConceptValue) -> ConceptValue:
    """Fold add-on rows (e.g. amortization of intangibles) into a value resolved
    via the trigger tag."""
    if c.add_when is None or cv.status != "OK" or cv.tag != f"us-gaap_{c.add_when[0]}":
        return cv
    extra = []
    for t in c.add_when[1]:
        hit = next((r for r in stmt.rows if r.key == f"us-gaap_{t}" and _usable(r, fy)), None)
        if hit is not None:
            extra.append(hit)
    if not extra:
        return cv
    return ConceptValue(
        c.name, fy, cv.value + c.sign * sum(r.values[fy] for r in extra), None,
        "SUM(" + "+".join([cv.tag] + [r.tag for r in extra]) + ")",
        " + ".join([cv.label] + [r.label for r in extra]), "DERIVED")


def _resolve_by_label(c: Concept, stmt: MergedStatement, fy: str) -> ConceptValue | None:
    if not c.label_re:
        return None
    label_re = re.compile(c.label_re, re.IGNORECASE)
    for r in stmt.rows:
        if _usable(r, fy) and _is_undimensioned(r) and label_re.search(r.label):
            return _ok(c, fy, r)
    return None


def _supplied_view(stmt: MergedStatement, fy: str) -> MergedStatement | None:
    """``stmt`` restricted to the rows whose ``fy`` value came from the year's source
    filing (``sources[fy]``), or None when provenance is unknown or nothing is filtered."""
    src = stmt.sources.get(fy)
    if src is None or not any(r.value_src for r in stmt.rows):
        return None
    rows = [r for r in stmt.rows if r.value_src.get(fy) == src.accession]
    return None if len(rows) == len(stmt.rows) else replace(stmt, rows=rows)


def _resolve_one(c: Concept, statements: dict[str, MergedStatement], fy: str) -> ConceptValue:
    """Tag matching across all of the concept's statements (in order) first, then
    label fallback across them; the first hit wins. Both passes run first on the rows the
    year's source filing supplied, so a year resolves from one filing: rows only another
    filing has (another entity's history after a predecessor recast, e.g. KNTK; an older
    presentation's cash rows, e.g. JPM/ENB FY2016) are used only when that finds nothing."""
    present = [(t, statements[t]) for t in c.stypes if t in statements]
    views = [(t, v) for t, s in present if (v := _supplied_view(s, fy)) is not None]
    for group in (views, present):
        for _, stmt in group:
            cv = _resolve_by_tag(c, stmt, fy)
            if cv is not None:
                return _add_on(c, stmt, fy, cv)
        for _, stmt in group:
            cv = _resolve_by_label(c, stmt, fy)
            if cv is not None:
                return cv
    return _unmapped(c.name, fy)


def resolve(statements: dict[str, MergedStatement]) -> dict[str, dict[str, ConceptValue]]:
    """concept -> fiscal year (oldest first) -> ConceptValue."""
    out: dict[str, dict[str, ConceptValue]] = {}
    fallback_years = next((s.fy_labels for s in statements.values()), [])
    for c in CONCEPTS:
        first = next((statements[t] for t in c.stypes if t in statements), None)
        years = first.fy_labels if first is not None else fallback_years
        out[c.name] = {fy: _resolve_one(c, statements, fy) for fy in years}
    _separate_cash_ends(out, statements)
    _reconcile_total_costs(out)
    return out


TOTAL_COST_TOLERANCE = 0.10


def _reconcile_total_costs(out: dict) -> None:
    """TotalCostsAndExpenses must be the IS total: within 10% of revenue - operating income.
    ``OperatingExpenses`` / ``OperatingCostsAndExpenses`` are often a single line or opex
    excluding cost of sales (TRGP "Operating expenses" 1,298 of 13,697; AAPL opex excluding
    COGS), which made the model's operating cost negative or far too small. Such a year is
    set UNMAPPED, so the Model derives operating cost from revenue - operating income."""
    for fy, tc in list(out.get("TotalCostsAndExpenses", {}).items()):
        rev = out.get("Revenue", {}).get(fy)
        oi = out.get("OperatingIncome", {}).get(fy)
        if tc.value is None or rev is None or oi is None or rev.value is None or oi.value is None:
            continue
        # Compared in absolute value: a negated-label total arrives negative here and is
        # flipped to expense-positive later by normalize_expense_signs (total costs are
        # never legitimately negative).
        implied = abs(rev.value - oi.value)
        if abs(abs(tc.value) - implied) > TOTAL_COST_TOLERANCE * max(implied, 1.0):
            out["TotalCostsAndExpenses"][fy] = _unmapped("TotalCostsAndExpenses", fy)


def _separate_cash_ends(out: dict, statements: dict[str, MergedStatement]) -> None:
    """Beginning and ending cash must be two rows. Where both resolved to one row, the row
    stays the ending balance and the beginning is re-resolved without it (remaining tags,
    then the positional rule); if nothing else is found the beginning is UNMAPPED, and
    ``run_checks`` reports the roll-forward as SKIP."""
    cf = statements.get("CF")
    begin = next(c for c in CONCEPTS if c.name == "CFBeginningCash")
    for fy, b in list(out.get("CFBeginningCash", {}).items()):
        e = out.get("CFEndingCash", {}).get(fy)
        if cf is None or e is None or b.row_key is None or b.row_key != e.row_key:
            continue
        view = replace(cf, rows=[r for r in cf.rows if r.key != b.row_key])
        out["CFBeginningCash"][fy] = _resolve_one(begin, {"CF": view}, fy)


def coverage_warnings(mapping: dict[str, dict[str, ConceptValue]]) -> list[str]:
    """One line per REQUIRED concept that is UNMAPPED in the latest year."""
    warnings = []
    for name in REQUIRED:
        by_fy = mapping.get(name) or {}
        if not by_fy:
            warnings.append(f"{name}: no data (statement missing)")
            continue
        fy = list(by_fy)[-1]
        if by_fy[fy].status == "UNMAPPED":
            warnings.append(f"{name}: UNMAPPED in {fy}")
    return warnings


# Model convention for these is "expense positive" (NIToNCI: positive = NCI's share of
# income), which is also the XBRL sign of their us-gaap tags. R-pages show the filer's
# presentation sign, so a negated-label row ("Interest expense (1,783)") arrives negative.
EXPENSE_CONCEPTS = ("TotalCostsAndExpenses", "DA", "Impairments", "InterestExpense",
                    "IncomeTax", "NIToNCI")
# Without an XBRL fact only these are flipped when negative: a tax benefit or a loss
# attributable to NCI is legitimately negative.
_NEVER_NEGATIVE = ("TotalCostsAndExpenses", "DA", "Impairments", "InterestExpense")
_STD_TAG = re.compile(r"us-gaap_\w+")


def normalize_expense_signs(mapping: dict[str, dict[str, ConceptValue]],
                            fact_values: Callable[[str], dict[str, float]]
                            ) -> dict[str, dict[str, ConceptValue]]:
    """Return ``mapping`` with the expense concepts in the model's "expense positive" sign.

    ``fact_values(tag)`` gives companyfacts ``{fy: value}`` for a bare us-gaap tag; it is
    the sign authority per (tag, year): a value whose sign differs from ``cv.sign * fact``
    is negated (and ``cv.sign`` flipped, so the Mapping sheet negates that cell). This uses
    only tags in the concept's own tag list. Custom tags, derived sums and listed tags with
    no fact for the year are flipped only when negative, and only for concepts that cannot
    be negative. A label-fallback hit on another us-gaap element is left as reported. The
    input mapping is not modified.
    """
    cache: dict[str, dict[str, float]] = {}

    def fact(tag: str, fy: str) -> float | None:
        if tag not in cache:
            cache[tag] = fact_values(tag) or {}
        return cache[tag].get(fy)

    concepts = {c.name: c for c in CONCEPTS}
    out = dict(mapping)
    for name in EXPENSE_CONCEPTS:
        if name not in mapping or name not in concepts:
            continue
        listed = set(concepts[name].tags)
        fixed = {}
        for fy, cv in mapping[name].items():
            flip = False
            if cv.value and cv.status != "UNMAPPED":
                std = bool(cv.tag and _STD_TAG.fullmatch(cv.tag))
                bare = (cv.tag or "").removeprefix("us-gaap_")
                f = fact(bare, fy) if std and bare in listed else None
                if f:
                    flip = (cv.value > 0) != (cv.sign * f > 0)
                elif not std or bare in listed:
                    flip = cv.value < 0 and name in _NEVER_NEGATIVE
                # else: a label-fallback hit on some other us-gaap element, whose XBRL sign
                # need not match the concept (KMI "(Gain) loss on divestitures and
                # impairments"): left as reported.
            fixed[fy] = replace(cv, value=-cv.value, sign=-cv.sign) if flip else cv
        out[name] = fixed
    return out
