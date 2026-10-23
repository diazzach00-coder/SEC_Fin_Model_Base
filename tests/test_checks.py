from datetime import date

import pytest

from sec_model.checks import run_checks, summarize
from sec_model.companyfacts import (
    dei_shares_outstanding, fill_from_facts, fy_values, load_companyfacts, unit_hint_from_facts,
)
from sec_model.types import ConceptValue, MergedRow, MergedStatement


def cv(concept, fy, value):
    if value is None:
        return ConceptValue(concept, fy, None, None, None, None, "UNMAPPED")
    return ConceptValue(concept, fy, value, concept, "us-gaap_" + concept, concept, "OK")


def mp(**by_concept):
    """mp(TotalAssets={"FY2024": 10000.0}) -> mapping"""
    return {c: {fy: cv(c, fy, v) for fy, v in d.items()} for c, d in by_concept.items()}


def result(results, name, fy="FY2024"):
    hits = [r for r in results if r.name == name and r.fy == fy]
    assert len(hits) == 1, [(r.name, r.fy) for r in results]
    return hits[0]


def mrow(label, values, indent=1, abstract=False, key=None):
    return MergedRow(key=key or label, tag="us-gaap_" + label.replace(" ", ""), label=label,
                     indent=indent, is_abstract=abstract, unit="money", values=values,
                     still_reported=True)


def stmt(stype, rows, years=("FY2024",)):
    return MergedStatement(stype=stype, currency="USD", fy_labels=list(years),
                           fy_ends={y: date(int(y[2:]), 12, 31) for y in years},
                           rows=rows, sources={})


def test_balance_sheet_pass_within_tolerance():
    r = run_checks({}, mp(TotalAssets={"FY2024": 10000.0},
                          LiabilitiesAndEquity={"FY2024": 10000.4}), {})
    assert result(r, "BS balances").status == "PASS"


def test_balance_sheet_fail():
    r = run_checks({}, mp(TotalAssets={"FY2024": 10000.0},
                          LiabilitiesAndEquity={"FY2024": 10100.0}), {})
    c = result(r, "BS balances")
    assert c.status == "FAIL"
    assert (c.lhs, c.rhs) == (10000.0, 10100.0)
    assert c.tolerance == pytest.approx(50.5)


def test_tolerance_floor_is_one_million():
    r = run_checks({}, mp(TotalAssets={"FY2024": 50.0}, LiabilitiesAndEquity={"FY2024": 50.9}), {})
    c = result(r, "BS balances")
    assert c.tolerance == 1.0 and c.status == "PASS"


def test_ni_tie_is_vs_cf():
    m = mp(NetIncome={"FY2023": 500.0, "FY2024": 600.0},
           CFNetIncome={"FY2023": 500.0, "FY2024": 590.0})
    r = run_checks({}, m, {})
    assert result(r, "NI ties IS=CF", "FY2023").status == "PASS"
    assert result(r, "NI ties IS=CF", "FY2024").status == "FAIL"


def test_cash_roll_forward():
    m = mp(CFBeginningCash={"FY2024": 100.0}, NetChangeInCash={"FY2024": 25.0},
           CFEndingCash={"FY2024": 125.0})
    c = result(run_checks({}, m, {}), "Cash roll-forward")
    assert c.status == "PASS" and c.lhs == 125.0 and c.rhs == 125.0
    m["CFEndingCash"]["FY2024"] = cv("CFEndingCash", "FY2024", 140.0)
    assert result(run_checks({}, m, {}), "Cash roll-forward").status == "FAIL"


def test_cf_cash_vs_bs_cash_is_warn_only():
    m = mp(CFEndingCash={"FY2024": 1000.0}, Cash={"FY2024": 800.0})  # restricted cash
    c = result(run_checks({}, m, {}), "CF cash vs BS cash")
    assert c.status == "WARN"
    m["Cash"]["FY2024"] = cv("Cash", "FY2024", 1000.0)
    assert result(run_checks({}, m, {}), "CF cash vs BS cash").status == "PASS"


def test_api_crosscheck_fail_above_5pct():
    m = mp(Revenue={"FY2022": 1000.0, "FY2023": 1000.0, "FY2024": 1000.0})
    facts = {"Revenue": {"FY2022": 1000.2, "FY2023": 1020.0, "FY2024": 1100.0}}
    r = run_checks({}, m, facts)
    assert result(r, "API cross-check: Revenue", "FY2022").status == "PASS"
    assert result(r, "API cross-check: Revenue", "FY2023").status == "WARN"
    assert result(r, "API cross-check: Revenue", "FY2024").status == "FAIL"


def test_unmapped_input_gives_skip():
    m = mp(TotalAssets={"FY2024": 10000.0}, LiabilitiesAndEquity={"FY2024": None})
    c = result(run_checks({}, m, {}), "BS balances")
    assert c.status == "SKIP" and "LiabilitiesAndEquity" in c.note
    # mapped, but the API value is absent
    api = result(run_checks({}, mp(Revenue={"FY2024": 5.0}), {}), "API cross-check: Revenue")
    assert api.status == "SKIP" and "companyfacts" in api.note


def test_subtotal_children_sum_warn():
    rows = [
        mrow("Assets", {}, indent=0, abstract=True),
        mrow("Cash", {"FY2024": 40.0}, indent=2),
        mrow("Receivables", {"FY2024": 60.0}, indent=2),
        mrow("Total current assets", {"FY2024": 100.0}, indent=1),
        mrow("Inventory", {"FY2024": 30.0}, indent=2),
        mrow("Other", {"FY2024": 10.0}, indent=2),
        mrow("Total other assets", {"FY2024": 45.0}, indent=1),
    ]
    r = run_checks({"BS": stmt("BS", rows)}, {}, {})
    ok = result(r, "Subtotal: Total current assets")
    assert ok.status == "PASS" and ok.lhs == 100.0 and ok.rhs == 100.0
    bad = result(r, "Subtotal: Total other assets")
    assert bad.status == "WARN" and bad.lhs == 45.0 and bad.rhs == 40.0


def test_subtotal_skips_none_child_and_flat_rows():
    rows = [
        mrow("A", {"FY2024": None}, indent=2),
        mrow("B", {"FY2024": 5.0}, indent=2),
        mrow("Total x", {"FY2024": 5.0}, indent=1),
        mrow("Total y", {"FY2024": 5.0}, indent=1),   # preceded only by an equal-indent row
    ]
    r = run_checks({"IS": stmt("IS", rows)}, {}, {})
    assert result(r, "Subtotal: Total x").status == "SKIP"
    flat = result(r, "Subtotal: Total y")
    assert flat.status == "SKIP" and flat.note == "no indented children"


def test_subtotal_stops_at_abstract_row():
    rows = [
        mrow("Old", {"FY2024": 999.0}, indent=2),
        mrow("Section", {}, indent=2, abstract=True),
        mrow("C1", {"FY2024": 3.0}, indent=2),
        mrow("Total z", {"FY2024": 3.0}, indent=1),
    ]
    c = result(run_checks({"BS": stmt("BS", rows)}, {}, {}), "Subtotal: Total z")
    assert c.status == "PASS"


def test_coverage_checks():
    m = mp(Revenue={"FY2024": 10.0}, Cash={"FY2024": None})
    sts = {"IS": stmt("IS", [mrow("Revenue", {"FY2024": 10.0})]),
           "BS": stmt("BS", [mrow("Cash", {"FY2024": 1.0})])}
    r = run_checks(sts, m, {})
    assert result(r, "Coverage: Revenue").status == "PASS"
    assert result(r, "Coverage: Cash").status == "WARN"
    assert result(r, "Coverage: Income statement").status == "PASS"
    assert result(r, "Coverage: Cash flow statement").status == "WARN"


def test_summarize_counts():
    m = mp(TotalAssets={"FY2024": 10.0}, LiabilitiesAndEquity={"FY2024": 10.0})
    results = run_checks({}, m, {})
    s = summarize(results)
    assert sum(s.values()) == len(results)
    assert s["PASS"] >= 1 and s["SKIP"] >= 1 and s["FAIL"] == 0 and set(s) == {"PASS", "WARN", "FAIL", "SKIP"}


# ---- companyfacts -------------------------------------------------------------------

def entry(end, val, filed, start=None, form="10-K", fp="FY", fy=None):
    e = {"end": end, "val": val, "fy": fy or int(end[:4]), "fp": fp, "form": form, "filed": filed}
    if start:
        e["start"] = start
    return e


def facts_with(tag, unit, entries, taxonomy="us-gaap"):
    return {"cik": 1, "facts": {taxonomy: {tag: {"units": {unit: entries}}}}}


def test_fy_values_picks_12_month_fy_entries_newest_filed():
    facts = facts_with("Revenues", "USD", [
        entry("2023-12-31", 100e6, "2024-02-20", start="2023-01-01"),
        entry("2023-12-31", 110e6, "2025-02-20", start="2023-01-01", fy=2024),  # restated, newer
        entry("2024-12-31", 200e6, "2025-02-20", start="2024-01-01"),
        entry("2024-09-30", 50e6, "2024-11-01", start="2024-07-01", form="10-Q", fp="Q3"),
        entry("2024-12-31", 60e6, "2025-02-20", start="2024-10-01"),           # 3-month
        entry("2022-12-31", 70e6, "2023-02-20", start="2022-01-01", form="8-K"),
    ])
    assert fy_values(facts, "Revenues", "USD") == {"FY2023": 110.0, "FY2024": 200.0}


def test_fy_values_instant_prefix_and_missing():
    facts = facts_with("Assets", "USD", [entry("2024-12-31", 5e9, "2025-02-20", form="10-K/A")])
    assert fy_values(facts, "Assets", "USD") == {"FY2024": 5000.0}
    assert fy_values(facts, "us-gaap_Assets", "USD") == {"FY2024": 5000.0}
    assert fy_values(facts, "Nope", "USD") == {}
    assert fy_values(facts, "Assets", "CAD") == {}


def test_fy_values_labels_by_end_date_not_fy_field():
    facts = facts_with("Revenues", "USD", [
        entry("2023-12-31", 100e6, "2025-02-20", start="2023-01-01", fy=2024)])
    assert fy_values(facts, "Revenues", "USD") == {"FY2023": 100.0}


def test_fy_values_per_share_unscaled():
    facts = facts_with("EarningsPerShareDiluted", "USD/shares",
                       [entry("2024-12-31", 2.5, "2025-02-20", start="2024-01-01")])
    assert fy_values(facts, "EarningsPerShareDiluted", "USD/shares") == {"FY2024": 2.5}


def test_dei_shares_outstanding():
    facts = facts_with("EntityCommonStockSharesOutstanding", "shares",
                       [entry("2025-01-31", 215_000_000, "2025-02-20")], taxonomy="dei")
    assert dei_shares_outstanding(facts) == {"FY2025": 215.0}
    assert dei_shares_outstanding({"facts": {}}) == {}


def test_unit_hint_from_facts():
    facts = {"facts": {"us-gaap": {
        "Revenues": {"units": {"USD": []}},
        "WeightedAverageNumberOfDilutedSharesOutstanding": {"units": {"shares": []}},
        "EarningsPerShareDiluted": {"units": {"USD/shares": []}},
        "Odd": {"units": {"pure": []}}}}}
    h = unit_hint_from_facts(facts)
    assert h("us-gaap_Revenues") == "money"
    assert h("us-gaap_WeightedAverageNumberOfDilutedSharesOutstanding") == "shares"
    assert h("us-gaap_EarningsPerShareDiluted") == "perunit"
    assert h("us-gaap_Odd") is None
    assert h("us-gaap_Missing") is None
    assert h("epd_Custom") is None


def test_load_companyfacts_url():
    class C:
        def get_json(self, url):
            self.url = url
            return {"ok": 1}
    c = C()
    assert load_companyfacts(c, 1039684) == {"ok": 1}
    assert c.url == "https://data.sec.gov/api/xbrl/companyfacts/CIK0001039684.json"


def test_fill_from_facts_fills_only_unmapped_years():
    facts = {"facts": {"us-gaap": {"WeightedAverageNumberOfDilutedSharesOutstanding": {"units": {
        "shares": [entry("2023-12-31", 99e6, "2024-02-20", start="2023-01-01"),
                   entry("2024-12-31", 101e6, "2025-02-20", start="2024-01-01")]}}}}}
    m = mp(DilutedUnits={"FY2023": None, "FY2024": 100.0}, Revenue={"FY2023": 1.0, "FY2024": 2.0})
    out = fill_from_facts(m, facts, "DilutedUnits",
                          ["WeightedAverageLimitedPartnershipUnitsOutstandingDiluted",
                           "WeightedAverageNumberOfDilutedSharesOutstanding"], "shares")
    f = out["FY2023"]
    assert (f.value, f.status, f.row_key, f.label) == (99.0, "DERIVED", None, "SEC companyfacts API")
    assert f.tag == "companyfacts:WeightedAverageNumberOfDilutedSharesOutstanding"
    assert out["FY2024"].value == 100.0 and out["FY2024"].status == "OK"
    assert m["DilutedUnits"]["FY2023"].status == "UNMAPPED"      # input not mutated


def test_fill_from_facts_dei_and_missing_concept():
    facts = facts_with("EntityCommonStockSharesOutstanding", "shares",
                       [entry("2024-12-31", 50e6, "2025-02-20")], taxonomy="dei")
    m = mp(Revenue={"FY2024": 1.0, "FY2025": 2.0})
    out = fill_from_facts(m, facts, "UnitsOutstanding",
                          ["dei:EntityCommonStockSharesOutstanding"], "shares")
    assert out["FY2024"].value == 50.0 and out["FY2024"].tag == \
        "companyfacts:dei:EntityCommonStockSharesOutstanding"
    assert "FY2025" not in out
