"""Resolve each model concept to one as-reported row per fiscal year."""
import re

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
    return ConceptValue(c.name, fy, c.sign * r.values[fy], r.key, r.tag, r.label, "OK")


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
    return ConceptValue(
        c.name, fy, c.sign * sum(r.values[fy] for r in used), None,
        "SUM(" + "+".join(r.tag for r in used) + ")",
        " + ".join(r.label for r in used), "DERIVED")


def _resolve_in(c: Concept, stmt: MergedStatement, fy: str) -> ConceptValue:
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

    if label_re:
        for r in rows:
            if _usable(r, fy) and _is_undimensioned(r) and label_re.search(r.label):
                return _ok(c, fy, r)
    return _unmapped(c.name, fy)


def _resolve_one(c: Concept, statements: dict[str, MergedStatement], fy: str) -> ConceptValue:
    """Try each of the concept's statements in order; the first hit wins."""
    for stype in c.stypes:
        stmt = statements.get(stype)
        if stmt is None:
            continue
        cv = _resolve_in(c, stmt, fy)
        if cv.status != "UNMAPPED":
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
    return out


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
