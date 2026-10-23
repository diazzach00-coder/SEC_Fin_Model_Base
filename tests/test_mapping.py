from dataclasses import replace
from datetime import date

import sec_model.mapping as mapping
from sec_model.concepts import CONCEPTS
from sec_model.mapping import REQUIRED, coverage_warnings, resolve
from sec_model.types import MergedRow, MergedStatement

CANONICAL = [
    "Revenue", "TotalCostsAndExpenses", "OperatingIncome", "DA", "Impairments", "Gains",
    "EquityEarnings", "InterestExpense", "PretaxIncome", "IncomeTax", "NetIncome", "NIToNCI",
    "DilutedUnits", "UnitsOutstanding", "OperatingCashFlow", "CFNetIncome", "DeferredTax",
    "DistributionsFromAffiliates", "Capex", "InvestingCF", "Distributions", "DistributionsToNCI",
    "FinancingCF", "NetChangeInCash", "CFBeginningCash", "CFEndingCash", "Cash", "Receivables",
    "Inventory", "NetPPE", "EquityInvestments", "TotalAssets", "Payables", "DebtCurrent",
    "LongTermDebt", "TotalDebt", "LiabilitiesAndEquity", "EquityCommon", "NCI",
    "TotalEquityInclNCI",
]


def mrow(key, label, values, tag=None, abstract=False):
    return MergedRow(key=key, tag=tag or key.split("|")[0], label=label, indent=0,
                     is_abstract=abstract, unit="money", values=values, still_reported=True)


def mstmt(stype, rows, years=("FY2023", "FY2024")):
    return MergedStatement(stype=stype, currency="USD", fy_labels=list(years),
                           fy_ends={y: date(int(y[2:]), 12, 31) for y in years},
                           rows=rows, sources={})


def test_every_canonical_concept_defined_once():
    names = [c.name for c in CONCEPTS]
    assert sorted(names) == sorted(CANONICAL)
    assert len(names) == len(set(names))


def test_concept_statement_scopes():
    by = {c.name: c.stype for c in CONCEPTS}
    cf = {"CFNetIncome", "DeferredTax", "DistributionsFromAffiliates", "Capex", "InvestingCF",
          "Distributions", "DistributionsToNCI", "FinancingCF", "NetChangeInCash",
          "CFBeginningCash", "CFEndingCash", "OperatingCashFlow"}
    bs = {"Cash", "Receivables", "Inventory", "NetPPE", "EquityInvestments", "TotalAssets",
          "Payables", "DebtCurrent", "LongTermDebt", "TotalDebt", "LiabilitiesAndEquity",
          "EquityCommon", "NCI", "TotalEquityInclNCI", "UnitsOutstanding"}
    for name, stype in by.items():
        assert stype == ("CF" if name in cf else "BS" if name in bs else "IS"), name


def test_tag_order_preferred_over_label():
    # "Total revenues" label would match the fallback, but the tag list wins.
    is_ = mstmt("IS", [
        mrow("us-gaap_Revenues", "Sales", {"FY2023": 10.0, "FY2024": 11.0}),
        mrow("us-gaap_SalesRevenueNet", "Total revenues", {"FY2023": 99.0, "FY2024": 99.0}),
        mrow("us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax", "Contract rev",
             {"FY2023": 55.0, "FY2024": 56.0}),
    ])
    out = resolve({"IS": is_})
    cv = out["Revenue"]["FY2024"]
    assert (cv.status, cv.value, cv.tag, cv.row_key) == ("OK", 11.0, "us-gaap_Revenues",
                                                         "us-gaap_Revenues")
    assert list(out["Revenue"]) == ["FY2023", "FY2024"]


def test_dimensional_rows_not_matched_by_tag_or_label():
    is_ = mstmt("IS", [
        mrow("us-gaap_Revenues|Product|srt_ProductOrServiceAxis=x", "Total revenues",
             {"FY2023": 5.0, "FY2024": 6.0}, tag="us-gaap_Revenues"),
        mrow("us-gaap_Revenues=Seg", "Segments", {"FY2023": None, "FY2024": None},
             tag="us-gaap_Revenues=Seg", abstract=True),
    ])
    assert resolve({"IS": is_})["Revenue"]["FY2024"].status == "UNMAPPED"


def test_abstract_and_none_values_skipped():
    is_ = mstmt("IS", [
        mrow("us-gaap_Revenues", "Revenues", {"FY2023": 1.0, "FY2024": 2.0}, abstract=True),
        mrow("us-gaap_SalesRevenueNet", "Net sales", {"FY2023": None, "FY2024": 7.0}),
    ])
    out = resolve({"IS": is_})["Revenue"]
    assert out["FY2023"].status == "UNMAPPED" and out["FY2023"].value is None
    assert out["FY2024"].tag == "us-gaap_SalesRevenueNet" and out["FY2024"].value == 7.0


def test_label_fallback_for_custom_tag():
    is_ = mstmt("IS", [
        mrow("us-gaap_Revenues", "Total revenues", {"FY2023": 100.0, "FY2024": 110.0}),
        mrow("epd_DepreciationAndAccretion", "Depreciation and accretion",
             {"FY2023": 4.0, "FY2024": 5.0}),
    ])
    cv = resolve({"IS": is_})["DA"]["FY2024"]
    assert (cv.status, cv.value, cv.tag, cv.row_key) == (
        "OK", 5.0, "epd_DepreciationAndAccretion", "epd_DepreciationAndAccretion")
    assert cv.label == "Depreciation and accretion"


def test_per_year_resolution_can_change_tag():
    is_ = mstmt("IS", [
        mrow("us-gaap_DepreciationDepletionAndAmortization", "D&A",
             {"FY2023": 3.0, "FY2024": None}),
        mrow("epd_DepreciationAndAccretion", "Depreciation and accretion",
             {"FY2023": None, "FY2024": 4.0}),
    ])
    da = resolve({"IS": is_})["DA"]
    assert da["FY2023"].status == da["FY2024"].status == "OK"
    assert da["FY2023"].tag == "us-gaap_DepreciationDepletionAndAmortization"
    assert da["FY2024"].tag == "epd_DepreciationAndAccretion"
    assert (da["FY2023"].value, da["FY2024"].value) == (3.0, 4.0)


def test_total_debt_sums_components():
    bs = mstmt("BS", [
        mrow("us-gaap_LongTermDebtNoncurrent", "Long-term debt", {"FY2023": 100.0, "FY2024": 120.0}),
        mrow("us-gaap_LongTermDebtCurrent", "Current maturities", {"FY2023": 10.0, "FY2024": 5.0}),
        mrow("us-gaap_CommercialPaper", "Commercial paper", {"FY2023": None, "FY2024": 7.0}),
        mrow("us-gaap_LongTermDebt", "Total debt", {"FY2023": 999.0, "FY2024": 999.0}),
    ])
    td = resolve({"BS": bs})["TotalDebt"]
    assert td["FY2023"].status == "DERIVED" and td["FY2023"].value == 110.0
    assert td["FY2024"].value == 132.0 and td["FY2024"].row_key is None
    assert td["FY2024"].tag.startswith("SUM(")
    assert "Commercial paper" in td["FY2024"].label


def test_total_debt_falls_back_to_long_term_debt():
    bs = mstmt("BS", [mrow("us-gaap_LongTermDebt", "Debt", {"FY2023": 50.0, "FY2024": 60.0})])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert (cv.status, cv.value, cv.tag) == ("OK", 60.0, "us-gaap_LongTermDebt")
    bs2 = mstmt("BS", [mrow("us-gaap_LongTermDebtAndCapitalLeaseObligations", "Debt",
                            {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"BS": bs2})["TotalDebt"]["FY2024"].value == 2.0


def test_unmapped_reported_and_required_flagged():
    is_ = mstmt("IS", [mrow("us-gaap_Revenues", "Revenues", {"FY2023": 1.0, "FY2024": 2.0})])
    out = resolve({"IS": is_})
    cv = out["NetIncome"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key, cv.tag, cv.label) == ("UNMAPPED", None, None,
                                                                    None, None)
    warns = coverage_warnings(out)
    assert any("NetIncome" in w for w in warns)
    flagged = {w.split()[0].rstrip(":") for w in warns}
    assert flagged == set(REQUIRED) - {"Revenue"}


def test_required_only_checks_latest_year():
    is_ = mstmt("IS", [mrow("us-gaap_Revenues", "Revenues", {"FY2023": None, "FY2024": 2.0})])
    assert not any(w.startswith("Revenue") for w in coverage_warnings(resolve({"IS": is_})))


def test_cf_beginning_vs_ending_cash_by_label():
    tag = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
    cf = mstmt("CF", [
        mrow(tag, "Cash at beginning of period", {"FY2023": 10.0, "FY2024": 15.0}),
        mrow(tag + "|Cash at end of period", "Cash at end of period",
             {"FY2023": 15.0, "FY2024": 22.0}, tag=tag),
    ])
    out = resolve({"CF": cf})
    assert out["CFBeginningCash"]["FY2024"].value == 15.0
    assert out["CFEndingCash"]["FY2024"].value == 22.0
    assert out["CFEndingCash"]["FY2024"].row_key == tag + "|Cash at end of period"


def test_cf_cash_swapped_row_order():
    tag = "us-gaap_CashAndCashEquivalentsAtCarryingValue"
    cf = mstmt("CF", [
        mrow(tag, "Cash and cash equivalents, ending balance", {"FY2023": 5.0, "FY2024": 8.0}),
        mrow(tag + "|Cash, beginning balance", "Cash, beginning balance",
             {"FY2023": 4.0, "FY2024": 5.0}, tag=tag),
    ])
    out = resolve({"CF": cf})
    assert out["CFBeginningCash"]["FY2024"].value == 5.0
    assert out["CFEndingCash"]["FY2024"].value == 8.0


def test_statement_scope_respected():
    # Cash tag present only on CF; BS "Cash" concept must not pick it up.
    cf = mstmt("CF", [mrow("us-gaap_CashAndCashEquivalentsAtCarryingValue", "Cash",
                           {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"CF": cf})["Cash"]["FY2024"].status == "UNMAPPED"


def test_missing_statement_flags_required_without_crash():
    warns = coverage_warnings(resolve({}))
    assert len(warns) == len(REQUIRED)


def test_income_tax_label_does_not_grab_pretax_row():
    is_ = mstmt("IS", [
        mrow("x_Pre", "Income before income taxes", {"FY2023": 9.0, "FY2024": 9.0}),
        mrow("x_Tax", "Provision for income taxes", {"FY2023": 2.0, "FY2024": 3.0}),
    ])
    assert resolve({"IS": is_})["IncomeTax"]["FY2024"].value == 3.0


def test_sign_applied(monkeypatch):
    neg = [replace(c, sign=-1) if c.name == "Revenue" else c for c in CONCEPTS]
    monkeypatch.setattr(mapping, "CONCEPTS", neg)
    is_ = mstmt("IS", [mrow("us-gaap_Revenues", "Revenues", {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"IS": is_})["Revenue"]["FY2024"].value == -2.0


def test_cf_cash_positional_fallback_when_labels_have_no_keyword():
    tag = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
    cf = mstmt("CF", [
        mrow(tag, "Cash, including restricted cash, January 1", {"FY2023": 1.0, "FY2024": 2.0}),
        mrow(tag + "|Cash, including restricted cash, December 31",
             "Cash, including restricted cash, December 31", {"FY2023": 2.0, "FY2024": 3.0},
             tag=tag),
    ])
    out = resolve({"CF": cf})
    assert out["CFBeginningCash"]["FY2024"].value == 2.0
    assert out["CFEndingCash"]["FY2024"].value == 3.0
