"""Thin wrapper over the SEC companyfacts API: unit hints, FY values, dei shares."""
from collections.abc import Callable
from datetime import date

from .types import ConceptValue

URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
FORMS = frozenset({"10-K", "10-K/A", "10-KT"})
_SCALED = 1e6


def load_companyfacts(client, cik: int) -> dict:
    return client.get_json(URL.format(cik=cik))


def _concept_block(facts: dict, tag: str) -> dict:
    """Look up ``tag`` in the facts. ``dei:Name`` reads the dei taxonomy; ``Name`` and
    ``us-gaap_Name`` read us-gaap."""
    all_facts = (facts or {}).get("facts", {})
    if tag.startswith("dei:"):
        return all_facts.get("dei", {}).get(tag[4:], {})
    return all_facts.get("us-gaap", {}).get(tag.removeprefix("us-gaap_"), {})


def unit_hint_from_facts(facts: dict) -> Callable[[str], str | None]:
    """Return ``tag -> "money" | "shares" | "perunit" | None`` for ``us-gaap_X`` tags."""
    gaap = (facts or {}).get("facts", {}).get("us-gaap", {})

    def hint(tag: str) -> str | None:
        if not tag.startswith("us-gaap_"):
            return None
        units = gaap.get(tag.removeprefix("us-gaap_"), {}).get("units", {})
        for key in units:
            if "/" in key:
                return "perunit"
            if key == "shares":
                return "shares"
            if len(key) == 3 and key.isalpha() and key.isupper():
                return "money"
        return None

    return hint


def fy_values(facts: dict, tag: str, currency: str) -> dict[str, float]:
    """``fy_label -> value`` for one tag, labelled by the entry's period end year.

    ``currency`` is the companyfacts unit key (``"USD"``, ``"CAD"``, ``"shares"``,
    ``"USD/shares"``). Money and share counts are returned in millions; per-share
    units are unscaled. Only annual (``fp == "FY"``) 10-K-family entries are used;
    duration entries must span 350-380 days, instant entries are kept as-is. Among
    duplicates for a year the newest ``filed`` wins.
    """
    entries = _concept_block(facts, tag).get("units", {}).get(currency)
    if not entries:
        return {}
    best: dict[str, tuple[str, float]] = {}
    for e in entries:
        if e.get("fp") != "FY" or e.get("form") not in FORMS:
            continue
        try:
            end = date.fromisoformat(e["end"])
            if e.get("start"):
                days = (end - date.fromisoformat(e["start"])).days
                if not 350 <= days <= 380:
                    continue
            val = float(e["val"])
        except (KeyError, TypeError, ValueError):
            continue
        fy = "FY" + str(end.year)
        filed = e.get("filed") or ""
        if fy not in best or filed >= best[fy][0]:
            best[fy] = (filed, val)
    scale = 1.0 if "/" in currency else 1 / _SCALED
    return {fy: v * scale for fy, (_, v) in sorted(best.items())}


def dei_shares_outstanding(facts: dict) -> dict[str, float]:
    """Cover-page shares outstanding in millions, labelled by the entry's end date year."""
    return fy_values(facts, "dei:EntityCommonStockSharesOutstanding", "shares")


def fill_from_facts(mapping: dict[str, dict[str, ConceptValue]], facts: dict, concept: str,
                    tags: list[str], unit: str) -> dict[str, ConceptValue]:
    """Return ``mapping[concept]`` with UNMAPPED/missing years filled from companyfacts.

    For each fiscal year in the mapping where the concept is UNMAPPED (or absent), the
    first tag in ``tags`` that has a value for the year supplies a DERIVED
    ``ConceptValue``. Years nothing can fill keep their existing entry (or stay absent).
    The input mapping is not modified.
    """
    existing = dict(mapping.get(concept) or {})
    years = sorted({fy for by_fy in mapping.values() for fy in by_fy} | set(existing))
    by_tag = {t: fy_values(facts, t, unit) for t in tags}
    out = dict(existing)
    for fy in years:
        cur = existing.get(fy)
        if cur is not None and cur.status != "UNMAPPED":
            continue
        for t in tags:
            if fy in by_tag[t]:
                out[fy] = ConceptValue(concept, fy, by_tag[t][fy], None,
                                       f"companyfacts:{t}", "SEC companyfacts API", "DERIVED")
                break
    return out
