from dataclasses import replace
from datetime import date
from pathlib import Path

import pytest

import sec_model.mapping as mapping
from sec_model.concepts import CONCEPTS
from sec_model.mapping import REQUIRED, coverage_warnings, resolve
from sec_model.merge import merge_statements
from sec_model.rfile import parse_rfile
from sec_model.types import Filing, MergedRow, MergedStatement

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
    by = {c.name: c.stypes for c in CONCEPTS}
    cf = {"CFNetIncome", "DeferredTax", "DistributionsFromAffiliates", "Capex", "InvestingCF",
          "Distributions", "DistributionsToNCI", "FinancingCF", "NetChangeInCash",
          "CFBeginningCash", "CFEndingCash", "OperatingCashFlow"}
    bs = {"Cash", "Receivables", "Inventory", "NetPPE", "EquityInvestments", "TotalAssets",
          "Payables", "DebtCurrent", "LongTermDebt", "TotalDebt", "LiabilitiesAndEquity",
          "EquityCommon", "NCI", "TotalEquityInclNCI", "UnitsOutstanding"}
    for name, stypes in by.items():
        expected = ("IS", "CF") if name == "DA" else (
            "CF" if name in cf else "BS" if name in bs else "IS",)
        assert stypes == expected, name


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


def test_da_resolves_from_cf_when_absent_from_is():
    is_ = mstmt("IS", [mrow("us-gaap_Revenues", "Revenues", {"FY2023": 1.0, "FY2024": 2.0})])
    cf = mstmt("CF", [mrow("epd_DepreciationAndAccretion", "Depreciation and accretion",
                           {"FY2023": 30.0, "FY2024": 40.0})])
    cv = resolve({"IS": is_, "CF": cf})["DA"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key) == ("OK", 40.0, "epd_DepreciationAndAccretion")


def test_da_prefers_is_over_cf():
    is_ = mstmt("IS", [mrow("us-gaap_DepreciationDepletionAndAmortization", "D&A",
                            {"FY2023": 1.0, "FY2024": 2.0})])
    cf = mstmt("CF", [mrow("us-gaap_DepreciationDepletionAndAmortization", "D&A",
                           {"FY2023": 8.0, "FY2024": 9.0})])
    assert resolve({"IS": is_, "CF": cf})["DA"]["FY2024"].value == 2.0


def test_da_cf_only_statement_set():
    cf = mstmt("CF", [mrow("us-gaap_Depreciation", "Depreciation", {"FY2023": 1.0, "FY2024": 2.0})])
    da = resolve({"CF": cf})["DA"]
    assert list(da) == ["FY2023", "FY2024"] and da["FY2024"].value == 2.0


def test_total_debt_orcl_style_notes_payable():
    bs = mstmt("BS", [
        mrow("us-gaap_NotesPayableCurrent", "Notes payable, current", {"FY2023": 7.0, "FY2024": 8.0}),
        mrow("us-gaap_LongTermNotesAndLoans", "Notes payable, non-current",
             {"FY2023": 85.0, "FY2024": 122.0}),
    ])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key) == ("DERIVED", 130.0, None)
    assert cv.tag == "SUM(us-gaap_LongTermNotesAndLoans+us-gaap_NotesPayableCurrent)"


def test_total_debt_oke_style():
    bs = mstmt("BS", [
        mrow("us-gaap_LongTermDebtCurrent", "Current maturities", {"FY2023": 1.0, "FY2024": 2.0}),
        mrow("us-gaap_ShortTermBorrowings", "Short-term borrowings", {"FY2023": 3.0, "FY2024": 4.0}),
        mrow("us-gaap_LongTermDebtNoncurrent", "LTD", {"FY2023": 10.0, "FY2024": 20.0}),
    ])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert (cv.status, cv.value) == ("DERIVED", 26.0)


def test_total_debt_first_present_within_group():
    bs = mstmt("BS", [
        mrow("us-gaap_LongTermDebtNoncurrent", "LTD", {"FY2023": 10.0, "FY2024": 20.0}),
        mrow("us-gaap_LongTermNotesAndLoans", "Notes", {"FY2023": 5.0, "FY2024": 5.0}),
    ])
    assert resolve({"BS": bs})["TotalDebt"]["FY2024"].value == 20.0


def test_total_debt_commercial_paper_not_double_counted_with_debt_current():
    bs = mstmt("BS", [
        mrow("us-gaap_LongTermDebtNoncurrent", "LTD", {"FY2023": 10.0, "FY2024": 20.0}),
        mrow("us-gaap_DebtCurrent", "Short-term debt", {"FY2023": 3.0, "FY2024": 4.0}),
        mrow("us-gaap_CommercialPaper", "Commercial paper", {"FY2023": 1.0, "FY2024": 2.0}),
    ])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert cv.value == 24.0
    assert "CommercialPaper" not in cv.tag


def test_total_debt_commercial_paper_added_without_debt_current():
    bs = mstmt("BS", [
        mrow("us-gaap_LongTermDebtNoncurrent", "LTD", {"FY2023": 10.0, "FY2024": 20.0}),
        mrow("us-gaap_LongTermDebtCurrent", "Current", {"FY2023": 3.0, "FY2024": 4.0}),
        mrow("us-gaap_CommercialPaper", "CP", {"FY2023": 1.0, "FY2024": 2.0}),
        mrow("us-gaap_ShortTermBorrowings", "STB", {"FY2023": 1.0, "FY2024": 1.0}),
    ])
    assert resolve({"BS": bs})["TotalDebt"]["FY2024"].value == 27.0


def test_total_debt_single_tag_fallbacks_and_unmapped():
    bs = mstmt("BS", [mrow("us-gaap_LongTermDebt", "Debt", {"FY2023": 50.0, "FY2024": 60.0})])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert (cv.value, cv.tag) == (60.0, "us-gaap_LongTermDebt")
    tag = "us-gaap_LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities"
    bs2 = mstmt("BS", [mrow(tag, "Debt", {"FY2023": 1.0, "FY2024": 2.0})])
    cv2 = resolve({"BS": bs2})["TotalDebt"]["FY2024"]
    assert (cv2.value, cv2.tag) == (2.0, tag)
    assert resolve({"BS": mstmt("BS", [])})["TotalDebt"]["FY2024"].status == "UNMAPPED"


def test_diluted_units_tags_and_label_fallback_ignores_eps():
    is_ = mstmt("IS", [
        mrow("us-gaap_NetIncomeLossNetOfTaxPerOutstandingLimitedPartnershipUnitDiluted",
             "Diluted earnings per common unit", {"FY2023": 2.0, "FY2024": 2.5}),
    ])
    assert resolve({"IS": is_})["DilutedUnits"]["FY2024"].status == "UNMAPPED"
    is2 = mstmt("IS", [
        mrow("epd_X", "Diluted earnings per unit", {"FY2023": 2.0, "FY2024": 2.5}),
        mrow("epd_WeightedDiluted", "Diluted weighted-average common units outstanding",
             {"FY2023": 2000.0, "FY2024": 2100.0}),
    ])
    cv = resolve({"IS": is2})["DilutedUnits"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key) == ("OK", 2100.0, "epd_WeightedDiluted")
    is3 = mstmt("IS", [mrow("us-gaap_WeightedAverageLimitedPartnershipUnitsOutstandingDiluted",
                            "Units", {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"IS": is3})["DilutedUnits"]["FY2024"].value == 2.0


# ---- fixture-backed smoke test ------------------------------------------------

FIXTURES = Path(__file__).parent / "fixtures"


def _load(co):
    statements = {}
    for st in ("IS", "BS", "CF"):
        ps = parse_rfile((FIXTURES / co / f"{st}.htm").read_text(encoding="utf-8"), st, "u")
        ends = sorted({c.end for c in ps.columns if c.months in (None, 12)})
        fy_ends = {f"FY{e.year}": e for e in ends}
        filing = Filing(1, "0000000001-26-000001", "10-K", date(2030, 1, 1), ends[-1], "x.htm")
        statements[st] = merge_statements([(filing, ps)], list(fy_ends), fy_ends)
    return resolve(statements)


NON_BANK = ["OKE", "EPD", "ET", "KNTK", "ORCL"]
SMOKE = ["Revenue", "NetIncome", "TotalAssets", "OperatingCashFlow", "Cash", "DA", "TotalDebt"]


@pytest.mark.parametrize("co", NON_BANK)
def test_fixture_smoke_non_bank(co):
    m = _load(co)
    for name in SMOKE:
        cv = list(m[name].values())[-1]
        assert cv.status in ("OK", "DERIVED") and cv.value is not None, (co, name)


def test_fixture_smoke_jpm():
    m = _load("JPM")
    for name in ("NetIncome", "TotalAssets", "Cash"):
        cv = list(m[name].values())[-1]
        assert cv.status in ("OK", "DERIVED") and cv.value is not None, name


def test_da_adds_intangible_amortization_when_resolved_via_bare_depreciation():
    cf = mstmt("CF", [
        mrow("us-gaap_Depreciation", "Depreciation", {"FY2023": 5.0, "FY2024": 6.0}),
        mrow("us-gaap_AmortizationOfIntangibleAssets", "Amortization of intangible assets",
             {"FY2023": None, "FY2024": 2.0}),
    ])
    da = resolve({"CF": cf})["DA"]
    cv = da["FY2024"]
    assert (cv.status, cv.value, cv.row_key) == ("DERIVED", 8.0, None)
    assert cv.tag == "SUM(us-gaap_Depreciation+us-gaap_AmortizationOfIntangibleAssets)"
    assert "Depreciation" in cv.label and "Amortization of intangible assets" in cv.label
    # No amortization value in FY2023: stays the plain row.
    assert (da["FY2023"].status, da["FY2023"].value) == ("OK", 5.0)


def test_da_does_not_add_amortization_to_a_total_da_tag():
    cf = mstmt("CF", [
        mrow("us-gaap_DepreciationDepletionAndAmortization", "D&A", {"FY2023": 5.0, "FY2024": 6.0}),
        mrow("us-gaap_AmortizationOfIntangibleAssets", "Amortization",
             {"FY2023": 1.0, "FY2024": 2.0}),
    ])
    cv = resolve({"CF": cf})["DA"]["FY2024"]
    assert (cv.status, cv.value) == ("OK", 6.0)


def test_tag_match_in_later_statement_beats_label_hit_in_earlier_statement():
    is_ = mstmt("IS", [mrow("x_Cost", "Cost of revenues (exclusive of depreciation)",
                            {"FY2023": 900.0, "FY2024": 950.0})])
    cf = mstmt("CF", [mrow("us-gaap_DepreciationDepletionAndAmortization", "Depreciation",
                           {"FY2023": 55.0, "FY2024": 60.0})])
    cv = resolve({"IS": is_, "CF": cf})["DA"]["FY2024"]
    assert (cv.value, cv.tag) == (60.0, "us-gaap_DepreciationDepletionAndAmortization")
    # With no tag anywhere, the label fallback still applies.
    cf2 = mstmt("CF", [mrow("epd_X", "Depreciation and accretion", {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"IS": is_, "CF": cf2})["DA"]["FY2024"].row_key == "x_Cost"


def test_fixture_da_tags():
    expect = {
        "OKE": "us-gaap_DepreciationDepletionAndAmortization",
        "EPD": "epd_DepreciationAndAccretion",
        "ORCL": "SUM(us-gaap_Depreciation+us-gaap_AmortizationOfIntangibleAssets)",
    }
    for co, tag in expect.items():
        cv = list(_load(co)["DA"].values())[-1]
        assert cv.tag == tag, co
    assert list(_load("ORCL")["DA"].values())[-1].status == "DERIVED"


def test_net_income_prefers_continuing_ops_incl_nci_over_controlling_interest():
    # ENB FY2016-FY2019: consolidated "Earnings" is tagged as continuing operations including
    # NCI (no ProfitLoss); NetIncomeLoss is "Earnings attributable to controlling interests".
    is_ = mstmt("IS", [
        mrow("us-gaap_IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest",
             "Earnings", {"FY2023": 2309.0, "FY2024": 3266.0}),
        mrow("us-gaap_NetIncomeLoss", "Earnings attributable to controlling interests",
             {"FY2023": 2069.0, "FY2024": 2859.0}),
    ])
    assert resolve({"IS": is_})["NetIncome"]["FY2024"].value == 3266.0
    both = mstmt("IS", [mrow("us-gaap_ProfitLoss", "Net income", {"FY2023": 1.0, "FY2024": 2.0})]
                 + is_.rows)
    assert resolve({"IS": both})["NetIncome"]["FY2024"].value == 2.0


def test_net_change_in_cash_restricted_basis_excluding_fx_tag():
    # MPLX reports the restricted-cash-basis net change with the "ExcludingExchangeRateEffect"
    # element; an older filing's cash-only net change must not be used in its place.
    base = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecrease"
    cf = mstmt("CF", [
        mrow("us-gaap_CashAndCashEquivalentsPeriodIncreaseDecrease",
             "Net increase (decrease) in cash and cash equivalents", {"FY2023": 191.0, "FY2024": None}),
        mrow(base + "ExcludingExchangeRateEffect",
             "Net change in cash, cash equivalents and restricted cash",
             {"FY2023": 187.0, "FY2024": 76.0}),
    ])
    out = resolve({"CF": cf})["NetChangeInCash"]
    assert out["FY2023"].value == 187.0 and out["FY2024"].value == 76.0


def test_nci_income_nonredeemable_tag_and_label_skips_including_portion_total():
    # ET FY2025: NetIncomeLossAttributableToNoncontrollingInterest is gone; the NCI share is
    # tagged "Nonredeemable". The pretax row labelled "Income (Loss), Including Portion
    # Attributable to Noncontrolling Interest" must never be taken as NCI income.
    is_ = mstmt("IS", [
        mrow("x_PretaxInclNCI", "Income (Loss), Including Portion Attributable to Noncontrolling Interest",
             {"FY2023": 7106.0, "FY2024": 6058.0}),
        mrow("us-gaap_NetIncomeLossAttributableToNonredeemableNoncontrollingInterest",
             "Less: Net income attributable to noncontrolling interests",
             {"FY2023": 1692.0, "FY2024": 1208.0}),
    ])
    assert resolve({"IS": is_})["NIToNCI"]["FY2024"].value == 1208.0
    label_only = mstmt("IS", [is_.rows[0],
                              mrow("x_NCI", "Less: Net income attributable to noncontrolling interests",
                                   {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"IS": label_only})["NIToNCI"]["FY2024"].value == 2.0


def test_pretax_income_newer_before_tax_element():
    # ET FY2022+: pretax income is IncomeLossIncludingPortionAttributableToNoncontrollingInterest
    # ("Income (Loss), Including Portion Attributable to Noncontrolling Interest, before Tax").
    is_ = mstmt("IS", [
        mrow("us-gaap_IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
             "Income from continuing operations before income tax expense",
             {"FY2023": None, "FY2024": None}),
        mrow("us-gaap_IncomeLossIncludingPortionAttributableToNoncontrollingInterest",
             "Income (Loss), Including Portion Attributable to Noncontrolling Interest",
             {"FY2023": 7106.0, "FY2024": 6058.0}),
        mrow("us-gaap_IncomeTaxExpenseBenefit", "Income tax expense", {"FY2023": 541.0, "FY2024": 350.0}),
    ])
    out = resolve({"IS": is_})
    assert out["PretaxIncome"]["FY2024"].value == 6058.0
    assert out["NIToNCI"]["FY2024"].status == "UNMAPPED"


# ---- expense sign normalisation (companyfacts is the sign authority) ----
def _facts(table):
    return lambda tag: table.get(tag, {})


def test_negated_label_expense_flipped_to_xbrl_sign():
    # OKE: "Interest expense (1,783)" and "Income taxes (1,028)" on the R-page; XBRL positive.
    is_ = mstmt("IS", [
        mrow("us-gaap_InterestExpenseNonoperating", "Interest expense", {"FY2023": -1371.0, "FY2024": -1783.0}),
        mrow("us-gaap_IncomeTaxExpenseBenefit", "Income taxes", {"FY2023": -998.0, "FY2024": -1028.0}),
    ])
    facts = _facts({"InterestExpenseNonoperating": {"FY2023": 1371.0, "FY2024": 1783.0},
                    "IncomeTaxExpenseBenefit": {"FY2023": 998.0, "FY2024": 1028.0}})
    out = mapping.normalize_expense_signs(resolve({"IS": is_}), facts)
    cv = out["InterestExpense"]["FY2024"]
    assert cv.value == 1783.0 and cv.sign == -1
    assert out["IncomeTax"]["FY2023"].value == 998.0 and out["IncomeTax"]["FY2023"].sign == -1


def test_tax_benefit_matching_fact_sign_is_kept():
    is_ = mstmt("IS", [mrow("us-gaap_IncomeTaxExpenseBenefit", "Income tax benefit",
                            {"FY2023": -50.0, "FY2024": 40.0})])
    facts = _facts({"IncomeTaxExpenseBenefit": {"FY2023": -50.0, "FY2024": 40.0}})
    out = mapping.normalize_expense_signs(resolve({"IS": is_}), facts)["IncomeTax"]
    assert (out["FY2023"].value, out["FY2023"].sign) == (-50.0, 1)
    assert (out["FY2024"].value, out["FY2024"].sign) == (40.0, 1)


def test_custom_tag_without_fact_flipped_only_for_never_negative_concepts():
    is_ = mstmt("IS", [
        mrow("x_InterestCost", "Interest expense", {"FY2023": -10.0, "FY2024": -12.0}),
        mrow("x_Taxes", "Provision for income taxes", {"FY2023": -3.0, "FY2024": -4.0}),
    ])
    out = mapping.normalize_expense_signs(resolve({"IS": is_}), _facts({}))
    assert out["InterestExpense"]["FY2024"].value == 12.0
    assert out["InterestExpense"]["FY2024"].sign == -1
    assert out["IncomeTax"]["FY2024"].value == -4.0       # a benefit is possible: left alone


def test_interest_net_tag_is_income_positive():
    # KMI: only "Interest, net" (InterestIncomeExpenseNonoperatingNet, XBRL -1,801 = expense).
    is_ = mstmt("IS", [mrow("us-gaap_InterestIncomeExpenseNonoperatingNet", "Interest, net",
                            {"FY2023": -1844.0, "FY2024": -1801.0})])
    res = resolve({"IS": is_})
    assert res["InterestExpense"]["FY2024"].value == 1801.0
    assert res["InterestExpense"]["FY2024"].sign == -1
    facts = _facts({"InterestIncomeExpenseNonoperatingNet": {"FY2023": -1844.0, "FY2024": -1801.0}})
    out = mapping.normalize_expense_signs(res, facts)["InterestExpense"]["FY2024"]
    assert (out.value, out.sign) == (1801.0, -1)


def test_normalisation_leaves_other_concepts_and_input_untouched():
    is_ = mstmt("IS", [mrow("us-gaap_Revenues", "Revenues", {"FY2023": -1.0, "FY2024": -2.0}),
                       mrow("us-gaap_InterestExpense", "Interest", {"FY2023": -1.0, "FY2024": -2.0})])
    res = resolve({"IS": is_})
    out = mapping.normalize_expense_signs(res, _facts({}))
    assert out["Revenue"]["FY2024"].value == -2.0
    assert res["InterestExpense"]["FY2024"].value == -2.0 and out["InterestExpense"]["FY2024"].value == 2.0


def test_label_fallback_on_other_us_gaap_element_left_as_reported():
    # KMI: Impairments falls back by label to GainLossOnSalesOfAssetsAndAssetImpairmentCharges,
    # a gain-positive element; -32 is a net gain and must not become a +32 impairment.
    is_ = mstmt("IS", [mrow("us-gaap_GainLossOnSalesOfAssetsAndAssetImpairmentCharges",
                            "(Gain) loss on divestitures and impairments, net",
                            {"FY2023": -32.0, "FY2024": -15.0})])
    facts = _facts({"GainLossOnSalesOfAssetsAndAssetImpairmentCharges": {"FY2023": 32.0, "FY2024": 15.0}})
    out = mapping.normalize_expense_signs(resolve({"IS": is_}), facts)["Impairments"]["FY2024"]
    assert (out.value, out.sign) == (-15.0, 1)


def test_mapping_prefers_rows_supplied_by_the_years_source_filing():
    # KNTK FY2021: its supplier (BCP's FY2022 10-K) reports revenue as RevenueFromContract...
    # and ending cash under one label; Altus's own FY2021 10-K (another entity before the
    # recast) supplies a "Revenues" row and a differently labelled ending-cash row.
    from sec_model.types import SourceRec
    tag = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
    new, old = "0001692787-23-000009", "0001784031-22-000008"
    fy = "FY2024"

    def r(key, label, value, accn, tag=None):
        x = mrow(key, label, {"FY2023": None, fy: value}, tag=tag)
        x.value_src = {fy: accn}
        return x
    is_ = mstmt("IS", [r("us-gaap_Revenues", "Total revenues", 160.6, old),
                       r("us-gaap_RevenueFromContractWithCustomerExcludingAssessedTax",
                         "Total operating revenues", 662.0, new)])
    cf = mstmt("CF", [r(tag, "CASH, BEGINNING OF PERIOD", 19.6, new),
                      r(tag + "|CASH AND CASH EQUIVALENTS AT END OF PERIOD",
                        "CASH AND CASH EQUIVALENTS AT END OF PERIOD", 132.0, old, tag=tag),
                      r(tag + "|CASH, END OF PERIOD", "CASH, END OF PERIOD", 18.7, new, tag=tag)])
    src = SourceRec(new, "10-K", None, "u", "")
    is_.sources, cf.sources = {fy: src}, {fy: src}
    out = resolve({"IS": is_, "CF": cf})
    assert out["Revenue"][fy].value == 662.0
    assert out["CFEndingCash"][fy].value == 18.7
    # Without provenance (no value_src) the tag order rule is unchanged.
    for x in is_.rows + cf.rows:
        x.value_src = {}
    assert resolve({"IS": is_, "CF": cf})["Revenue"][fy].value == 160.6


def test_beginning_cash_never_shares_the_ending_cash_row():
    # A single cash row whose label matches both ends: it stays the ending balance and the
    # beginning is re-resolved from the remaining rows (here: none -> UNMAPPED).
    tag = "us-gaap_CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
    cf = mstmt("CF", [mrow(tag, "Cash at beginning and end of period", {"FY2023": 9.0, "FY2024": 15.0})])
    out = resolve({"CF": cf})
    assert out["CFEndingCash"]["FY2024"].row_key == tag
    assert out["CFBeginningCash"]["FY2024"].status == "UNMAPPED"
    other = "us-gaap_CashAndCashEquivalentsAtCarryingValue"
    cf2 = mstmt("CF", cf.rows + [mrow(other, "Cash at beginning of period", {"FY2023": 1.0, "FY2024": 9.0})])
    assert resolve({"CF": cf2})["CFBeginningCash"]["FY2024"].row_key == other


def test_supplier_label_match_beats_tag_match_in_another_filings_row():
    # JPM FY2016: its supplier (FY2017 10-K) reports cash as "Cash and due from banks ... at
    # the end of the period" (CashAndDueFromBanks, not a CF cash tag); the FY2018 10-K adds a
    # CashAndCashEquivalentsAtCarryingValue row on a broader basis. The year must resolve
    # from its supplier (label fallback) rather than mix in the other filing's row.
    from sec_model.types import SourceRec
    sup, other, fy = "0000019617-18-000057", "0000019617-19-000054", "FY2024"

    def r(key, label, value, accn, tag=None):
        x = mrow(key, label, {"FY2023": None, fy: value}, tag=tag)
        x.value_src = {fy: accn}
        return x
    cf = mstmt("CF", [
        r("us-gaap_CashAndDueFromBanks", "Cash and due from banks at the beginning of the period", 20490.0, sup),
        r("us-gaap_CashAndDueFromBanks|Cash and due from banks at the end of the period",
          "Cash and due from banks at the end of the period", 23873.0, sup,
          tag="us-gaap_CashAndDueFromBanks"),
        r("us-gaap_CashAndCashEquivalentsAtCarryingValue",
          "Cash and deposits with banks at the beginning of the period", 361683.0, other),
        r("us-gaap_CashAndCashEquivalentsAtCarryingValue|Cash and deposits with banks at the end of the period",
          "Cash and deposits with banks at the end of the period", 391154.0, other,
          tag="us-gaap_CashAndCashEquivalentsAtCarryingValue"),
    ])
    cf.sources = {fy: SourceRec(sup, "10-K", None, "u", "")}
    out = resolve({"CF": cf})
    assert out["CFBeginningCash"][fy].value == 20490.0
    assert out["CFEndingCash"][fy].value == 23873.0


def test_total_debt_long_term_including_current_plus_short_term_borrowings():
    # CVX/JPM: "Short-term debt" (ShortTermBorrowings) + "Long-term debt" tagged
    # ...IncludingCurrentMaturities. Total debt was the short-term line alone (977 vs 40,758).
    tag = "us-gaap_LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities"
    bs = mstmt("BS", [
        mrow("us-gaap_ShortTermBorrowings", "Short-term debt", {"FY2023": 4406.0, "FY2024": 977.0}),
        mrow(tag, "Long-term debt", {"FY2023": 20135.0, "FY2024": 39781.0}),
        mrow("us-gaap_LongTermDebtCurrent", "Current maturities (already included)",
             {"FY2023": 1.0, "FY2024": 2.0}),
    ])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert cv.value == 40758.0 and cv.status == "DERIVED"
    assert "LongTermDebtCurrent" not in cv.tag


def test_interest_expense_operating_before_net_interest_income():
    is_ = mstmt("IS", [
        mrow("us-gaap_InterestExpenseOperating", "Interest expense", {"FY2023": 101350.0, "FY2024": 97898.0}),
        mrow("us-gaap_InterestIncomeExpenseNet", "Net interest income", {"FY2023": 92583.0, "FY2024": 95443.0}),
    ])
    cv = resolve({"IS": is_})["InterestExpense"]["FY2024"]
    assert (cv.value, cv.tag) == (97898.0, "us-gaap_InterestExpenseOperating")


def test_total_costs_that_do_not_reconcile_to_revenue_less_operating_income_are_dropped():
    # TRGP: "Operating expenses" (OperatingCostsAndExpenses) is one line, not the total.
    is_ = mstmt("IS", [
        mrow("us-gaap_Revenues", "Total revenues", {"FY2023": 16381.5, "FY2024": 17028.3}),
        mrow("us-gaap_OperatingCostsAndExpenses", "Operating expenses", {"FY2023": 1175.6, "FY2024": 1298.3}),
        mrow("us-gaap_OperatingIncomeLoss", "Income from operations", {"FY2023": 2695.4, "FY2024": 3331.2}),
    ])
    assert resolve({"IS": is_})["TotalCostsAndExpenses"]["FY2024"].status == "UNMAPPED"
    # EPD-style: total costs differ from revenue - operating income only by equity earnings.
    epd = mstmt("IS", [
        mrow("us-gaap_Revenues", "Revenues", {"FY2023": 56219.0, "FY2024": 52596.0}),
        mrow("us-gaap_CostsAndExpenses", "Total costs and expenses", {"FY2023": 49289.0, "FY2024": 45691.0}),
        mrow("us-gaap_OperatingIncomeLoss", "Operating income", {"FY2023": 7338.0, "FY2024": 7266.0}),
    ])
    assert resolve({"IS": epd})["TotalCostsAndExpenses"]["FY2024"].value == 45691.0


def test_interest_and_debt_expense_maps_to_interest_expense():
    # MPLX: "Net interest and other financial costs" (InterestAndDebtExpense); it was UNMAPPED,
    # giving a 0% default interest rate and a 0% pre-tax cost of debt on the WACC sheet.
    is_ = mstmt("IS", [mrow("us-gaap_InterestAndDebtExpense", "Net interest and other financial costs",
                            {"FY2023": 921.0, "FY2024": 983.0})])
    assert resolve({"IS": is_})["InterestExpense"]["FY2024"].value == 983.0


def test_negated_total_costs_that_reconcile_are_kept_and_flipped_positive():
    # A total-costs row presented negative ("(45,691)") must survive the revenue - operating
    # income reconciliation and then be normalised to expense-positive.
    is_ = mstmt("IS", [
        mrow("us-gaap_Revenues", "Revenues", {"FY2023": 56219.0, "FY2024": 52596.0}),
        mrow("us-gaap_CostsAndExpenses", "Total costs and expenses", {"FY2023": -49289.0, "FY2024": -45691.0}),
        mrow("us-gaap_OperatingIncomeLoss", "Operating income", {"FY2023": 7338.0, "FY2024": 7266.0}),
    ])
    res = resolve({"IS": is_})
    assert res["TotalCostsAndExpenses"]["FY2024"].value == -45691.0
    out = mapping.normalize_expense_signs(res, _facts({"CostsAndExpenses": {"FY2024": 45691.0}}))
    cv = out["TotalCostsAndExpenses"]["FY2024"]
    assert (cv.value, cv.sign, cv.status) == (45691.0, -1, "OK")
    assert out["TotalCostsAndExpenses"]["FY2023"].value == 49289.0      # no fact: never negative
    bad = mstmt("IS", [is_.rows[0], is_.rows[2],
                       mrow("us-gaap_OperatingCostsAndExpenses", "Operating expenses",
                            {"FY2023": -1175.6, "FY2024": -1298.3})])
    assert resolve({"IS": bad})["TotalCostsAndExpenses"]["FY2024"].status == "UNMAPPED"


def test_total_debt_single_component_sum_keeps_the_row_link():
    # Only one sum-group component present -> status OK with that row's key and tag
    # (formula-linked on the Mapping sheet), not a hard-coded DERIVED value.
    bs = mstmt("BS", [mrow("us-gaap_LongTermDebtNoncurrent", "Long-term debt",
                           {"FY2023": 100.0, "FY2024": 120.0})])
    cv = resolve({"BS": bs})["TotalDebt"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key, cv.tag) == (
        "OK", 120.0, "us-gaap_LongTermDebtNoncurrent", "us-gaap_LongTermDebtNoncurrent")


def test_trgp_shaped_distributions_skip_the_nci_row():
    # TRGP: both payout rows are custom tags, so both concepts resolve by label. The NCI
    # row came first and was taken as common distributions; "Contributions from
    # noncontrolling interests" (an inflow) was taken as distributions to NCI.
    cf = mstmt("CF", [
        mrow("us-gaap_ProceedsFromMinorityShareholders",
             "Contributions from noncontrolling interests", {"FY2024": 12.0, "FY2025": 5.0}),
        mrow("trgp_PaymentsOfDividendsToNoncontrollingInterests",
             "Distributions to noncontrolling interests", {"FY2024": -232.6, "FY2025": -33.5}),
        mrow("trgp_PaymentsOfOrdinaryDividendsCommonAndPreferredShareholders",
             "Dividends paid to common shareholders", {"FY2024": -615.5, "FY2025": -818.3}),
    ], years=("FY2024", "FY2025"))
    out = resolve({"CF": cf})
    d, n = out["Distributions"]["FY2025"], out["DistributionsToNCI"]["FY2025"]
    assert (d.value, d.row_key) == (
        -818.3, "trgp_PaymentsOfOrdinaryDividendsCommonAndPreferredShareholders")
    assert (n.value, n.row_key) == (-33.5, "trgp_PaymentsOfDividendsToNoncontrollingInterests")
    assert out["DistributionsToNCI"]["FY2024"].value == -232.6


def test_distributions_and_nci_never_share_one_row():
    # A us-gaap distributions tag on an NCI-labelled row: the NCI label fallback hits the
    # same row. It stays with DistributionsToNCI; Distributions is UNMAPPED.
    cf = mstmt("CF", [mrow("us-gaap_PaymentsOfDividends",
                           "Distributions to noncontrolling interests",
                           {"FY2023": -5.0, "FY2024": -6.0})])
    out = resolve({"CF": cf})
    assert out["Distributions"]["FY2024"].status == "UNMAPPED"
    assert out["DistributionsToNCI"]["FY2024"].value == -6.0
    # The reverse: an NCI tag on a row labelled as common distributions stays common.
    cf = mstmt("CF", [mrow("us-gaap_PaymentsOfDividendsMinorityInterest",
                           "Distributions paid to unitholders",
                           {"FY2023": -5.0, "FY2024": -6.0})])
    out = resolve({"CF": cf})
    assert out["Distributions"]["FY2024"].value == -6.0
    assert out["DistributionsToNCI"]["FY2024"].status == "UNMAPPED"
