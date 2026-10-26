"""Entity-aware words: a corporation has shares and dividends, a partnership units and
distributions. Every such word in the workbook comes from `LABELS` via `label(id, entity)`;
the model's mechanics are identical for both."""
from __future__ import annotations

ENTITY_TYPES = ("corporation", "partnership")

LABELS: dict[str, dict[str, str]] = {
    "units": {"corporation": "Diluted shares", "partnership": "Diluted units"},
    "units_out": {"corporation": "Shares outstanding", "partnership": "Units outstanding"},
    "per_unit": {"corporation": "per share", "partnership": "per unit"},
    "dist": {"corporation": "Dividends paid", "partnership": "Distributions paid"},
    "dpu": {"corporation": "Dividend per share", "partnership": "Distribution per unit"},
    "dpu_growth": {"corporation": "Dividend per share growth",
                   "partnership": "Distribution per unit growth"},
    "ni_common": {"corporation": "Net income to common", "partnership": "Net income to partners"},
    "equity": {"corporation": "Shareholders' equity", "partnership": "Partners' capital"},
    "holder": {"corporation": "shareholders", "partnership": "unitholders"},
    "eps": {"corporation": "EPS (diluted)", "partnership": "Earnings per unit (diluted)"},
    "dist_coverage": {"corporation": "Dividend coverage (OCF - capex) / dividends",
                      "partnership": "Distribution coverage (OCF - capex) / distributions"},
}


def label(label_id: str, entity_type: str) -> str:
    """The text of ``label_id`` for ``entity_type`` ("corporation" | "partnership")."""
    if entity_type not in ENTITY_TYPES:
        raise ValueError(f"entity_type must be one of {ENTITY_TYPES}, got {entity_type!r}")
    return LABELS[label_id][entity_type]


def unit_noun(entity_type: str) -> str:
    """"share" / "unit": the noun in ``label("per_unit")``, for "Share price", "Basic
    shares (MM)" and the like."""
    return label("per_unit", entity_type).removeprefix("per ")


_LP_MARKERS = ("Partnership", "LimitedPartner")


def _tag(mapping: dict, concept: str, fy: str) -> str:
    cv = (mapping.get(concept) or {}).get(fy)
    return (cv.tag if cv is not None else None) or ""


def detect_entity_type(mapping: dict, latest_fy: str) -> str:
    """"partnership" when the latest year's common-equity tag is ``us-gaap_PartnersCapital``
    or its diluted-units or distributions tag names a (limited) partnership; else
    "corporation"."""
    if _tag(mapping, "EquityCommon", latest_fy) == "us-gaap_PartnersCapital":
        return "partnership"
    for concept in ("DilutedUnits", "Distributions"):
        tag = _tag(mapping, concept, latest_fy)
        if any(m in tag for m in _LP_MARKERS):
            return "partnership"
    return "corporation"
