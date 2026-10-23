"""Tie-out checks run on every build (spec section 5.5)."""
from .types import CheckResult, ConceptValue, MergedStatement

API_CONCEPTS = ("Revenue", "NetIncome", "TotalAssets", "OperatingCashFlow")
COVERAGE_STATEMENTS = (("IS", "Income statement"), ("BS", "Balance sheet"),
                       ("CF", "Cash flow statement"))
COVERAGE_ANCHORS = ("Revenue", "OperatingIncome", "DA", "Cash", "TotalDebt")
STATUSES = ("PASS", "WARN", "FAIL", "SKIP")
FAIL_REL_DIFF = 0.05


def _tol(lhs: float, rhs: float) -> float:
    return max(1.0, 0.005 * max(abs(lhs), abs(rhs)))


def _get(mapping: dict, concept: str, fy: str) -> float | None:
    cv: ConceptValue | None = (mapping.get(concept) or {}).get(fy)
    if cv is None or cv.status == "UNMAPPED" or cv.value is None:
        return None
    return cv.value


def _skip(name: str, fy: str, note: str) -> CheckResult:
    return CheckResult(name, fy, "SKIP", None, None, 0.0, note)


def _compare(name: str, fy: str, lhs: float, rhs: float, bad: str, note: str = "") -> CheckResult:
    tol = _tol(lhs, rhs)
    return CheckResult(name, fy, "PASS" if abs(lhs - rhs) <= tol else bad, lhs, rhs, tol, note)


def _years(statements: dict, mapping: dict, facts_values: dict) -> list[str]:
    years = {fy for s in statements.values() for fy in s.fy_labels}
    if not years:
        years = {fy for by_fy in mapping.values() for fy in by_fy}
    return sorted(years)


def _missing(mapping: dict, fy: str, *concepts: str) -> list[str]:
    return [c for c in concepts if _get(mapping, c, fy) is None]


def _tie_check(name: str, fy: str, mapping: dict, lhs_concepts: tuple[str, ...],
               rhs_concept: str, bad: str, note: str = "") -> CheckResult:
    gone = _missing(mapping, fy, *lhs_concepts, rhs_concept)
    if gone:
        return _skip(name, fy, "missing input: " + ", ".join(gone))
    lhs = sum(_get(mapping, c, fy) for c in lhs_concepts)
    return _compare(name, fy, lhs, _get(mapping, rhs_concept, fy), bad, note)


def _api_check(concept: str, fy: str, mapping: dict, facts_values: dict) -> CheckResult:
    name = f"API cross-check: {concept}"
    ours = _get(mapping, concept, fy)
    if ours is None:
        return _skip(name, fy, f"missing input: {concept}")
    theirs = (facts_values.get(concept) or {}).get(fy)
    if theirs is None:
        return _skip(name, fy, "no companyfacts value for this year")
    tol = _tol(ours, theirs)
    diff = abs(ours - theirs)
    if diff <= tol:
        return CheckResult(name, fy, "PASS", ours, theirs, tol, "")
    rel = diff / max(abs(ours), abs(theirs))
    status = "FAIL" if rel > FAIL_REL_DIFF else "WARN"
    return CheckResult(name, fy, status, ours, theirs, tol, f"differs by {rel:.1%}")


def _subtotal_checks(stmt: MergedStatement) -> list[CheckResult]:
    out: list[CheckResult] = []
    rows = stmt.rows
    for i, total in enumerate(rows):
        if total.is_abstract or not total.label.strip().lower().startswith("total"):
            continue
        children = []
        for prev in reversed(rows[:i]):
            if prev.is_abstract or prev.indent <= total.indent:
                break
            children.append(prev)
        name = f"Subtotal: {total.label.strip()}"
        for fy in stmt.fy_labels:
            tv = total.values.get(fy)
            if tv is None:
                continue
            if not children:
                out.append(_skip(name, fy, "no indented children"))
                continue
            vals = [c.values.get(fy) for c in children]
            if any(v is None for v in vals):
                out.append(_skip(name, fy, "a child value is missing"))
                continue
            out.append(_compare(name, fy, tv, sum(vals), "WARN"))
    return out


def _statement_has(stmt: MergedStatement | None, fy: str) -> bool:
    return (stmt is not None and fy in stmt.fy_labels and
            any(not r.is_abstract and r.values.get(fy) is not None for r in stmt.rows))


def run_checks(statements: dict[str, MergedStatement], mapping: dict[str, dict[str, ConceptValue]],
               facts_values: dict[str, dict[str, float]]) -> list[CheckResult]:
    years = _years(statements, mapping, facts_values)
    results: list[CheckResult] = []
    for fy in years:
        results.append(_tie_check("BS balances", fy, mapping, ("TotalAssets",),
                                  "LiabilitiesAndEquity", "FAIL"))
        results.append(_tie_check("NI ties IS=CF", fy, mapping, ("NetIncome",),
                                  "CFNetIncome", "FAIL"))
        results.append(_tie_check("Cash roll-forward", fy, mapping,
                                  ("CFBeginningCash", "NetChangeInCash"), "CFEndingCash", "FAIL"))
        results.append(_tie_check("CF cash vs BS cash", fy, mapping, ("CFEndingCash",), "Cash",
                                  "WARN", "restricted cash is often included in CF cash"))
        for concept in API_CONCEPTS:
            results.append(_api_check(concept, fy, mapping, facts_values))
    for stype in ("IS", "BS", "CF"):
        if stype in statements:
            results.extend(_subtotal_checks(statements[stype]))
    for fy in years:
        for stype, title in COVERAGE_STATEMENTS:
            ok = _statement_has(statements.get(stype), fy)
            results.append(CheckResult(f"Coverage: {title}", fy, "PASS" if ok else "WARN",
                                       None, None, 0.0, "" if ok else "statement missing"))
        for concept in COVERAGE_ANCHORS:
            ok = _get(mapping, concept, fy) is not None
            results.append(CheckResult(f"Coverage: {concept}", fy, "PASS" if ok else "WARN",
                                       None, None, 0.0, "" if ok else "not mapped"))
    return results


def summarize(results: list[CheckResult]) -> dict[str, int]:
    counts = dict.fromkeys(STATUSES, 0)
    for r in results:
        counts[r.status] = counts.get(r.status, 0) + 1
    return counts
