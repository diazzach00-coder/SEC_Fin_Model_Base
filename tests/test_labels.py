"""Entity-aware labels (shares vs units) and the no-sector-words guard."""
import re
from pathlib import Path

import pytest

from sec_model.labels import ENTITY_TYPES, LABELS, detect_entity_type, label
from sec_model.model_spec import build_spec
from sec_model.types import ConceptValue

ROOT = Path(__file__).resolve().parent.parent
FY = "FY2025"


def _cv(concept, tag, fy=FY):
    return ConceptValue(concept, fy, 1.0, "k", tag, "label", "OK")


def _mapping(**tags):
    """{concept: {fy: ConceptValue}} with one latest-year tag per concept."""
    return {c: {FY: _cv(c, t)} for c, t in tags.items()}


def test_label_table_shape():
    assert ENTITY_TYPES == ("corporation", "partnership")
    for lid, texts in LABELS.items():
        assert set(texts) == set(ENTITY_TYPES), lid
    assert (label("units", "corporation"), label("units", "partnership")) == \
        ("Diluted shares", "Diluted units")
    assert (label("dpu", "corporation"), label("dpu", "partnership")) == \
        ("Dividend per share", "Distribution per unit")
    assert (label("equity", "corporation"), label("equity", "partnership")) == \
        ("Shareholders' equity", "Partners' capital")
    assert (label("holder", "corporation"), label("holder", "partnership")) == \
        ("shareholders", "unitholders")
    with pytest.raises(ValueError):
        label("units", "trust")


def test_detect_entity_type_partnership_and_corporation():
    # EPD-shaped: partners' capital, LP unit and LP distribution tags
    epd = _mapping(EquityCommon="us-gaap_PartnersCapital",
                   DilutedUnits="us-gaap_WeightedAverageLimitedPartnershipUnitsOutstandingDiluted",
                   Distributions="us-gaap_DistributionMadeToLimitedPartnerCashDistributionsPaid")
    assert detect_entity_type(epd, FY) == "partnership"
    # any one signal is enough
    assert detect_entity_type(_mapping(EquityCommon="us-gaap_PartnersCapital"), FY) == \
        "partnership"
    assert detect_entity_type(_mapping(
        DilutedUnits="companyfacts:WeightedAverageLimitedPartnershipUnitsOutstandingDiluted"),
        FY) == "partnership"
    assert detect_entity_type(_mapping(
        Distributions="us-gaap_DistributionMadeToLimitedPartnerCashDistributionsPaid"),
        FY) == "partnership"
    # OKE-shaped corporation
    oke = _mapping(EquityCommon="us-gaap_StockholdersEquity",
                   DilutedUnits="us-gaap_WeightedAverageNumberOfDilutedSharesOutstanding",
                   Distributions="us-gaap_PaymentsOfDividendsCommonStock")
    assert detect_entity_type(oke, FY) == "corporation"
    # only the latest year counts; missing / unmapped values fall back to corporation
    old = {"EquityCommon": {"FY2020": _cv("EquityCommon", "us-gaap_PartnersCapital", "FY2020"),
                            FY: _cv("EquityCommon", "us-gaap_StockholdersEquity")}}
    assert detect_entity_type(old, FY) == "corporation"
    assert detect_entity_type({}, FY) == "corporation"
    assert detect_entity_type({"EquityCommon": {FY: None}}, FY) == "corporation"
    assert detect_entity_type(_mapping(EquityCommon=None), FY) == "corporation"


def test_spec_labels_follow_entity_type():
    epd, oke = build_spec("ppe", "partnership"), build_spec("ppe", "corporation")
    assert build_spec("ppe").line("UNITS").label == "Diluted shares"   # default: corporation
    for spec, e in ((epd, "partnership"), (oke, "corporation")):
        assert spec.entity_type == e
        assert spec.line("UNITS").label == label("units", e)
        assert spec.line("DPU_FCST").label == label("dpu", e)
        assert spec.line("DPU_HIST").label == label("dpu", e)
        assert spec.line("DIST").label == label("dist", e)
        assert spec.line("NI_COMMON").label == label("ni_common", e)
        assert spec.line("EQUITY").label == label("equity", e)
        assert spec.line("EPS").label == label("eps", e)
        drivers = {d.key: d.label for d in spec.drivers}
        assert drivers["DPU_GROWTH"] == label("dpu_growth", e)
        assert drivers["UNITS"] == label("units", e)
    assert epd.line("UNITS").label == "Diluted units"
    assert oke.line("UNITS").label == "Diluted shares"
    assert (oke.line("EPS").label, epd.line("EPS").label) == \
        ("EPS (diluted)", "Earnings per unit (diluted)")
    # keys and formulas are identical: only words change
    assert epd.line_keys() == oke.line_keys()
    assert [d.key for d in epd.drivers] == [d.key for d in oke.drivers]
    with pytest.raises(ValueError):
        build_spec("ppe", "trust")


def test_jv_toggle_defaults_off():
    jv = next(s for s in build_spec("ppe").scalars if s.key == "JV_TOGGLE")
    assert jv.default == 0.0
    assert jv.note == ("1 = include equity-method earnings in EBITDA (common for partnerships "
                       "and JV-heavy companies)")


_SECTOR = re.compile(r"midstream|east daley|asset inputs", re.I)
_PIPELINE_IDENT = re.compile(r"from \.pipeline\b|\bpipeline\.|\bpipeline import\b"
                             r"|import pipeline\b|sec_model\.pipeline\b")
_PIPELINE_WORD = re.compile(r"\bpipelines?\b", re.I)


def _scanned():
    return sorted((ROOT / "sec_model").rglob("*.py")) + [ROOT / "app.py",
                                                       ROOT / "scripts" / "regress.py"]


def test_no_sector_words():
    hits = []
    for path in _scanned():
        rel = path.relative_to(ROOT).as_posix()
        for n, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            if _SECTOR.search(line):
                hits.append(f"{rel}:{n}: {line.strip()}")
            if rel != "sec_model/pipeline.py" and \
                    _PIPELINE_WORD.search(_PIPELINE_IDENT.sub("", line)):
                hits.append(f"{rel}:{n}: {line.strip()}")
    assert not hits, "\n".join(hits)
