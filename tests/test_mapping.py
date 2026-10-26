from dataclasses import replace
from datetime import date
from pathlib import Path

import pytest

import sec_model.mapping as mapping
from sec_model.concepts import CONCEPTS
from sec_model.forecast import compute
from sec_model.mapping import REQUIRED, coverage_warnings, resolve
from sec_model.merge import merge_statements
from sec_model.model_spec import build_spec
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
    "TotalEquityInclNCI", "RevolverDrawn",
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
          "EquityCommon", "NCI", "TotalEquityInclNCI", "UnitsOutstanding", "RevolverDrawn"}
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


def test_revolver_drawn_tags_label_fallback_and_unmapped():
    c = next(c for c in CONCEPTS if c.name == "RevolverDrawn")
    assert c.tags == ("LineOfCredit", "LongTermLineOfCredit", "LinesOfCreditCurrent",
                      "RevolvingCreditFacility")
    assert c.stypes == ("BS",) and c.sum_groups == ()
    for tag in c.tags:
        bs = mstmt("BS", [mrow(f"us-gaap_{tag}", "Borrowings", {"FY2023": 3.0, "FY2024": 4.0})])
        cv = resolve({"BS": bs})["RevolverDrawn"]["FY2024"]
        assert (cv.status, cv.value, cv.tag) == ("OK", 4.0, f"us-gaap_{tag}"), tag
    bs = mstmt("BS", [mrow("xyz_Revolver", "Revolving credit facility",
                           {"FY2023": 5.0, "FY2024": 6.0}),
                      mrow("us-gaap_LongTermDebtNoncurrent", "Long-term debt",
                           {"FY2023": 100.0, "FY2024": 100.0})])
    cv = resolve({"BS": bs})["RevolverDrawn"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key) == ("OK", 6.0, "xyz_Revolver")
    bs2 = mstmt("BS", [mrow("xyz_Loc", "Borrowings under line of credit",
                            {"FY2023": 1.0, "FY2024": 2.0}),
                       mrow("us-gaap_LongTermDebt", "Debt", {"FY2023": 50.0, "FY2024": 50.0})])
    assert resolve({"BS": bs2})["RevolverDrawn"]["FY2024"].value == 2.0
    empty = resolve({"BS": mstmt("BS", [mrow("us-gaap_LongTermDebt", "Debt",
                                             {"FY2023": 1.0, "FY2024": 2.0})])})
    assert empty["RevolverDrawn"]["FY2024"].status == "UNMAPPED"
    # BS scope only: a cash-flow revolver line is not a drawn balance
    cf = mstmt("CF", [mrow("us-gaap_ProceedsFromLinesOfCredit", "Revolver borrowings",
                           {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve({"CF": cf})["RevolverDrawn"]["FY2024"].status == "UNMAPPED"


@pytest.mark.parametrize("label", [
    "Long-term debt, including revolving credit facility",
    "Total debt and line of credit borrowings",
    "Senior notes and revolver",
    "Notes payable under line of credit",
])
def test_revolver_label_fallback_skips_total_and_term_debt_rows(label):
    bs = mstmt("BS", [mrow("xyz_Debt", label, {"FY2023": 5.0, "FY2024": 6.0}),
                      mrow("us-gaap_LongTermDebt", "Debt", {"FY2023": 90.0, "FY2024": 90.0})])
    assert resolve({"BS": bs})["RevolverDrawn"]["FY2024"].status == "UNMAPPED"


@pytest.mark.parametrize("rev", [100.0, 120.0])
def test_revolver_at_or_above_total_debt_is_unmapped(rev):
    bs = mstmt("BS", [mrow("us-gaap_LongTermDebt", "Debt", {"FY2023": 100.0, "FY2024": 100.0}),
                      mrow("xyz_Rev", "Revolving credit facility",
                           {"FY2023": 40.0, "FY2024": rev})])
    rv = resolve({"BS": bs})["RevolverDrawn"]
    assert rv["FY2023"].value == 40.0                    # below total: kept (per year)
    assert rv["FY2024"].status == "UNMAPPED" and rv["FY2024"].value is None


def test_revolver_unmapped_when_total_debt_unmapped():
    bs = mstmt("BS", [mrow("xyz_Rev", "Revolving credit facility",
                           {"FY2023": 40.0, "FY2024": 50.0})])
    out = resolve({"BS": bs})
    assert out["TotalDebt"]["FY2024"].status == "UNMAPPED"
    assert out["RevolverDrawn"]["FY2024"].status == "UNMAPPED"


def _debt_split(rows):
    """Resolve a BS and run the model: (mapping, model values) for FY2023-FY2024."""
    out = resolve({"BS": mstmt("BS", rows)})
    hist = {c: {fy: cv.value for fy, cv in yrs.items()} for c, yrs in out.items()}
    return out, compute(build_spec("ppe"), hist, ["FY2023", "FY2024"], ["FY2025E"])


def test_separately_presented_revolver_joins_component_total_debt():
    out, m = _debt_split([
        mrow("us-gaap_LongTermDebtNoncurrent", "Long-term debt", {"FY2023": 200.0, "FY2024": 200.0}),
        mrow("us-gaap_LongTermDebtCurrent", "Current maturities", {"FY2023": 10.0, "FY2024": 10.0}),
        mrow("us-gaap_LineOfCredit", "Credit facility", {"FY2023": 50.0, "FY2024": 50.0}),
    ])
    td = out["TotalDebt"]["FY2024"]
    assert td.value == 260.0 and "us-gaap_LineOfCredit" in td.tag
    assert out["RevolverDrawn"]["FY2024"].value == 50.0
    assert (m["DEBT"]["FY2024"], m["REVOLVER"]["FY2024"]) == (210.0, 50.0)


def test_revolver_assumed_inside_a_single_total_including_tag():
    for tag in ("LongTermDebt", "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities"):
        out, m = _debt_split([
            mrow(f"us-gaap_{tag}", "Total debt", {"FY2023": 110.0, "FY2024": 110.0}),
            mrow("us-gaap_LineOfCredit", "Credit facility", {"FY2023": 30.0, "FY2024": 30.0}),
        ])
        assert out["TotalDebt"]["FY2024"].value == 110.0, tag
        assert out["RevolverDrawn"]["FY2024"].value == 30.0, tag
        assert (m["DEBT"]["FY2024"], m["REVOLVER"]["FY2024"]) == (80.0, 30.0), tag


def test_custom_revolver_outside_a_component_built_total_is_unmapped():
    out, m = _debt_split([
        mrow("us-gaap_LongTermDebtNoncurrent", "Long-term debt", {"FY2023": 200.0, "FY2024": 200.0}),
        mrow("us-gaap_LongTermDebtCurrent", "Current maturities", {"FY2023": 10.0, "FY2024": 10.0}),
        mrow("xyz_Revolver", "Revolving credit facility", {"FY2023": 50.0, "FY2024": 50.0}),
    ])
    td = out["TotalDebt"]["FY2024"]
    assert td.value == 210.0 and td.tag.startswith("SUM(") and "xyz_Revolver" not in td.tag
    assert out["RevolverDrawn"]["FY2024"].status == "UNMAPPED"
    assert (m["DEBT"]["FY2024"], m["REVOLVER"]["FY2024"]) == (210.0, 0.0)


def test_loc_tag_inside_a_component_built_total_is_kept():
    out, _ = _debt_split([
        mrow("us-gaap_LongTermDebtNoncurrent", "Long-term debt", {"FY2023": 200.0, "FY2024": 200.0}),
        mrow("us-gaap_LongTermLineOfCredit", "Credit facility", {"FY2023": 40.0, "FY2024": 40.0}),
    ])
    td = out["TotalDebt"]["FY2024"]
    assert td.value == 240.0 and "us-gaap_LongTermLineOfCredit" in td.tag
    rv = out["RevolverDrawn"]["FY2024"]
    assert (rv.status, rv.value) == ("OK", 40.0)


def test_custom_revolver_with_single_total_including_tag_is_kept_as_inside():
    # Current behaviour: a single total-including tag is not a SUM, so the separately
    # presented face row is assumed to be inside it and is split out.
    out, m = _debt_split([
        mrow("us-gaap_LongTermDebt", "Total debt", {"FY2023": 110.0, "FY2024": 110.0}),
        mrow("xyz_Revolver", "Revolving credit facility", {"FY2023": 30.0, "FY2024": 30.0}),
    ])
    assert out["TotalDebt"]["FY2024"].tag == "us-gaap_LongTermDebt"
    assert out["RevolverDrawn"]["FY2024"].value == 30.0
    assert (m["DEBT"]["FY2024"], m["REVOLVER"]["FY2024"]) == (80.0, 30.0)


def test_revolver_only_balance_sheet_has_no_negative_term_debt():
    out, m = _debt_split([
        mrow("us-gaap_LineOfCredit", "Credit facility", {"FY2023": 50.0, "FY2024": 50.0}),
    ])
    assert out["TotalDebt"]["FY2024"].value == 50.0
    assert out["RevolverDrawn"]["FY2024"].value == 50.0
    assert (m["DEBT"]["FY2024"], m["REVOLVER"]["FY2024"]) == (0.0, 50.0)


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


def _hdr(axis, member, label):
    return mrow(f"{axis}={member}", label, {"FY2023": None, "FY2024": None},
                tag=f"{axis}={member}", abstract=True)


_RP_AXIS = "us-gaap_RelatedPartyTransactionsByRelatedPartyAxis"


def test_receivables_reported_only_by_member_resolve_as_member_sum():
    # EPD-shaped: AR and AP appear only under the related-party axis.
    ar, ap = "us-gaap_AccountsReceivableNetCurrent", "us-gaap_AccountsPayableCurrent"
    bs = mstmt("BS", [
        mrow(ar, "Accounts receivable - trade", {"FY2023": None, "FY2024": None}),
        _hdr(_RP_AXIS, "us-gaap_NonrelatedPartyMember", "Nonrelated Party"),
        mrow(f"{ar}|Accounts receivable, net|us-gaap_NonrelatedPartyMember",
             "Accounts receivable, net", {"FY2023": 9236.0, "FY2024": 6494.0}, tag=ar),
        mrow(f"{ap}|Accounts payable|us-gaap_NonrelatedPartyMember", "Accounts payable",
             {"FY2023": 1227.0, "FY2024": 1021.0}, tag=ap),
        _hdr(_RP_AXIS, "us-gaap_RelatedPartyMember", "Related Party"),
        mrow(f"{ar}|Accounts receivable, net|us-gaap_RelatedPartyMember",
             "Accounts receivable, net", {"FY2023": 4.0, "FY2024": 1.0}, tag=ar),
        mrow(f"{ap}|Accounts payable|us-gaap_RelatedPartyMember", "Accounts payable",
             {"FY2023": 198.0, "FY2024": 217.0}, tag=ap),
    ])
    out = resolve({"BS": bs})
    cv = out["Receivables"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key, cv.tag) == (
        "DERIVED", 6495.0, None, f"SUM({ar} members)")
    assert "NonrelatedPartyMember" in cv.label and "RelatedPartyMember" in cv.label
    assert out["Payables"]["FY2024"].value == 1238.0
    assert out["Payables"]["FY2023"].value == 1425.0
    assert coverage_warnings(out, {"BS": bs}) == [
        w for w in coverage_warnings(out) if not w.startswith(("Receivables", "Payables"))]


def test_member_rows_ignored_when_the_tag_has_a_consolidated_row_or_axes_differ():
    ar = "us-gaap_AccountsReceivableNetCurrent"
    member = mrow(f"{ar}|AR|us-gaap_RelatedPartyMember", "AR", {"FY2023": 4.0, "FY2024": 1.0},
                  tag=ar)
    hdr = _hdr(_RP_AXIS, "us-gaap_RelatedPartyMember", "Related Party")
    total = mrow(ar, "Accounts receivable", {"FY2023": 50.0, "FY2024": 60.0})
    assert resolve({"BS": mstmt("BS", [total, hdr, member])})["Receivables"]["FY2024"].value \
        == 60.0
    # members of two different axes overlap: no sum
    other = mrow(f"{ar}|AR|srt_ProductMember", "AR", {"FY2023": 3.0, "FY2024": 3.0}, tag=ar)
    bs = mstmt("BS", [hdr, member, _hdr("srt_ProductOrServiceAxis", "srt_ProductMember", "P"),
                      other])
    assert resolve({"BS": bs})["Receivables"]["FY2024"].status == "UNMAPPED"
    # one member: kept linked to its row
    cv = resolve({"BS": mstmt("BS", [hdr, member])})["Receivables"]["FY2024"]
    assert (cv.status, cv.value, cv.row_key) == ("OK", 1.0, member.key)


def test_enb_shaped_affiliate_row_skipped_for_real_receivables_and_payables():
    bs = mstmt("BS", [
        mrow("enb_AccountsReceivablesAndOtherAmountsNetCurrent",
             "Accounts receivable and other", {"FY2023": None, "FY2024": None}),
        mrow("us-gaap_AccountsAndOtherReceivablesNetCurrent",
             "Trade receivables and unbilled revenues", {"FY2023": 6920.0, "FY2024": 7081.0}),
        mrow("us-gaap_OtherReceivablesNetCurrent", "Accounts receivable from affiliates",
             {"FY2023": 90.0, "FY2024": None}),
        mrow("us-gaap_AccountsPayableAndOtherAccruedLiabilitiesCurrent",
             "Trade payables and accrued liabilities", {"FY2023": 7060.0, "FY2024": 7555.0}),
        mrow("us-gaap_OtherAccountsPayableAndAccruedLiabilities",
             "Accounts payable to affiliates", {"FY2023": 22.0, "FY2024": None}),
    ])
    out = resolve({"BS": bs})
    for fy, ar, ap in (("FY2023", 6920.0, 7060.0), ("FY2024", 7081.0, 7555.0)):
        assert out["Receivables"][fy].value == ar
        assert out["Payables"][fy].value == ap
    # the label fallback no longer takes affiliate / related-party rows
    bs = mstmt("BS", [mrow("x_DueFromAffiliates", "Accounts receivable from affiliates",
                           {"FY2023": 90.0, "FY2024": 86.0}),
                      mrow("x_DueToRelated", "Accounts payable - related parties",
                           {"FY2023": 22.0, "FY2024": 38.0})])
    out = resolve({"BS": bs})
    assert out["Receivables"]["FY2024"].status == "UNMAPPED"
    assert out["Payables"]["FY2024"].status == "UNMAPPED"


def test_coverage_warns_when_working_capital_unmapped_but_bs_shows_it():
    bs = mstmt("BS", [mrow("x_DueFromAffiliates", "Accounts receivable from affiliates",
                           {"FY2023": 90.0, "FY2024": 86.0}),
                      mrow("x_Abs", "Payables:", {"FY2023": None, "FY2024": None},
                           abstract=True)])
    out = resolve({"BS": bs})
    warnings = coverage_warnings(out, {"BS": bs})
    assert any(w.startswith("Receivables: UNMAPPED in FY2024") for w in warnings)
    assert not any(w.startswith("Payables") for w in warnings)   # abstract row only
    assert not any(w.startswith(("Receivables", "Payables")) for w in coverage_warnings(out))


def test_fixture_epd_working_capital_resolves_from_members():
    m = _load("EPD")
    ar, ap = m["Receivables"]["FY2025"], m["Payables"]["FY2025"]
    assert (ar.status, ar.value) == ("DERIVED", 6495.0)
    assert (ap.status, ap.value) == ("DERIVED", 1238.0)


def test_member_sum_only_for_money_rows():
    # MPLX: units outstanding by class (common + Series B preferred) must not be added.
    tag = "us-gaap_LimitedPartnersCapitalAccountUnitsOutstanding"
    axis = "us-gaap_StatementClassOfStockAxis"
    rows = [_hdr(axis, "us-gaap_CommonMember", "Common"),
            mrow(f"{tag}|Units|us-gaap_CommonMember", "Units", {"FY2023": 600.0, "FY2024": 600.0},
                 tag=tag),
            _hdr(axis, "us-gaap_SeriesBPreferredStockMember", "Series B"),
            mrow(f"{tag}|Units|us-gaap_SeriesBPreferredStockMember", "Units",
                 {"FY2023": 0.6, "FY2024": 0.6}, tag=tag)]
    for r in rows:
        r.unit = "shares" if not r.is_abstract else r.unit
    out = resolve({"BS": mstmt("BS", rows)})
    assert out["UnitsOutstanding"]["FY2024"].status == "UNMAPPED"


def test_cvx_accounts_and_notes_receivable_tag():
    bs = mstmt("BS", [mrow("us-gaap_AccountsNotesAndLoansReceivableNetCurrent",
                           "Accounts and notes receivable", {"FY2023": 20684.0, "FY2024": 18075.0})])
    assert resolve({"BS": bs})["Receivables"]["FY2024"].value == 18075.0


def test_pretax_and_net_income_are_required():
    assert {"PretaxIncome", "NetIncome"} <= set(REQUIRED)
    is_ = mstmt("IS", [mrow("us-gaap_Revenues", "Revenue", {"FY2023": 1.0, "FY2024": 2.0})])
    assert "PretaxIncome: UNMAPPED in FY2024" in coverage_warnings(resolve({"IS": is_}))


def _kmi_shaped(equity_earnings: bool):
    is_rows = [mrow("us-gaap_Revenues", "Revenues", {"FY2023": 15000.0, "FY2024": 16000.0})]
    if equity_earnings:
        is_rows.append(mrow("us-gaap_IncomeLossFromEquityMethodInvestments",
                            "Earnings from equity investments", {"FY2023": 840.0, "FY2024": 896.0}))
    bs = mstmt("BS", [mrow("us-gaap_LongTermInvestments", "Investments",
                           {"FY2023": 7845.0, "FY2024": 7532.0})])
    return {"IS": mstmt("IS", is_rows), "BS": bs}


def test_long_term_investments_are_equity_investments_when_equity_earnings_mapped():
    cv = resolve(_kmi_shaped(True))["EquityInvestments"]["FY2024"]
    assert (cv.status, cv.value, cv.tag) == ("OK", 7532.0, "us-gaap_LongTermInvestments")


def test_long_term_investments_unmapped_without_equity_earnings():
    out = resolve(_kmi_shaped(False))
    assert out["EquityInvestments"]["FY2024"].status == "UNMAPPED"
    # the "^investments$" label alternative is gated too
    st = _kmi_shaped(False)
    st["BS"] = mstmt("BS", [mrow("x_Investments", "Investments", {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve(st)["EquityInvestments"]["FY2024"].status == "UNMAPPED"
    st = _kmi_shaped(True)
    st["BS"] = mstmt("BS", [mrow("x_Investments", "Investments", {"FY2023": 1.0, "FY2024": 2.0})])
    assert resolve(st)["EquityInvestments"]["FY2024"].value == 2.0
    # an ungated tag is still used without equity earnings
    st = _kmi_shaped(False)
    st["BS"] = mstmt("BS", [mrow("us-gaap_EquityMethodInvestments", "Equity investments",
                                 {"FY2023": 5.0, "FY2024": 6.0}),
                            mrow("us-gaap_LongTermInvestments", "Investments",
                                 {"FY2023": 7.0, "FY2024": 8.0})])
    assert resolve(st)["EquityInvestments"]["FY2024"].value == 6.0


# --- Review Focus 5: a lone member row does not pre-empt the label fallback -----------------

def test_lone_related_party_member_loses_to_custom_trade_receivables_row():
    ar = "us-gaap_AccountsReceivableNetCurrent"
    hdr = _hdr(_RP_AXIS, "us-gaap_RelatedPartyMember", "Related Party")
    member = mrow(f"{ar}|Accounts receivable|us-gaap_RelatedPartyMember",
                  "Accounts receivable", {"FY2023": 4.0, "FY2024": 1.0}, tag=ar)
    trade = mrow("x_TradeReceivablesNetCurrent", "Trade receivables, net",
                 {"FY2023": 900.0, "FY2024": 950.0})
    for rows in ([hdr, member, trade], [trade, hdr, member]):
        cv = resolve({"BS": mstmt("BS", rows)})["Receivables"]["FY2024"]
        assert (cv.status, cv.value, cv.row_key) == ("OK", 950.0, trade.key)
    # two or more members (EPD: related + nonrelated) still pre-empt the label fallback
    hdr2 = _hdr(_RP_AXIS, "us-gaap_NonrelatedPartyMember", "Nonrelated Party")
    nonrel = mrow(f"{ar}|Accounts receivable|us-gaap_NonrelatedPartyMember",
                  "Accounts receivable", {"FY2023": 800.0, "FY2024": 850.0}, tag=ar)
    cv = resolve({"BS": mstmt("BS", [hdr2, nonrel, hdr, member, trade])})["Receivables"]["FY2024"]
    assert (cv.status, cv.value) == ("DERIVED", 851.0)


@pytest.mark.parametrize("label, warns", [
    ("Income taxes payable", False), ("Accrued interest payable", False),
    ("Interest receivable", False), ("Notes receivable - interest bearing", False),
    ("Other receivables", False), ("Accounts payable", True), ("Trade payables", True),
    ("Notes receivable", True), ("Receivables", True), ("Accounts receivable, net", True),
])
def test_coverage_working_capital_regex(label, warns):
    bs = mstmt("BS", [mrow("x_Row", label, {"FY2023": 1.0, "FY2024": 2.0})])
    out = {"Receivables": {"FY2024": mapping._unmapped("Receivables", "FY2024")},
           "Payables": {"FY2024": mapping._unmapped("Payables", "FY2024")}}
    got = [w for w in coverage_warnings(out, {"BS": bs})
           if w.startswith(("Receivables", "Payables"))]
    assert bool(got) is warns, got


@pytest.mark.parametrize("tag", ["us-gaap_LongTermInvestments", "us-gaap_Investments"])
@pytest.mark.parametrize("label", ["Marketable securities", "Long-term investments in securities",
                                   "Available-for-sale investments"])
def test_gated_investment_tags_skip_securities_rows(tag, label):
    st = _kmi_shaped(True)
    st["BS"] = mstmt("BS", [mrow(tag, label, {"FY2023": 70.0, "FY2024": 80.0})])
    assert resolve(st)["EquityInvestments"]["FY2024"].status == "UNMAPPED"
    # an equity-method row under the same gated tag is still taken
    st["BS"] = mstmt("BS", [mrow(tag, "Investments", {"FY2023": 70.0, "FY2024": 80.0})])
    assert resolve(st)["EquityInvestments"]["FY2024"].value == 80.0
