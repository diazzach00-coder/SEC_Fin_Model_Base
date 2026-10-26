import math

import pytest

from sec_model.formula import Fn, Line, RenderContext, to_excel
from sec_model.forecast import (
    compute, driver_defaults, eval_order, iter_nodes, scalar_defaults, validate_spec,
)
from sec_model.model_spec import LineDef, ModelSpec, build_spec

FY = ["FY2023", "FY2024", "FY2025"]
FCST = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]


def _h(*vals):
    return dict(zip(FY, vals))


# A toy company with round numbers. Statements tie: OperatingIncome = Revenue - TCE,
# NetIncome = Pretax - Tax, LiabilitiesAndEquity = TotalAssets.
TOY = {
    "Revenue": _h(1000, 1100, 1200),
    "TotalCostsAndExpenses": _h(800, 870, 940),
    "OperatingIncome": _h(200, 230, 260),
    "DA": _h(100, 110, 120),
    "Impairments": _h(0, 10, 0),
    "Gains": _h(0, 0, 5),
    "EquityEarnings": _h(50, 50, 60),
    "InterestExpense": _h(40, 42, 44),
    "PretaxIncome": _h(210, 238, 286),
    "IncomeTax": _h(10, 12, 14),
    "NetIncome": _h(200, 226, 272),
    "NIToNCI": _h(20, 20, 20),
    "DilutedUnits": _h(100, 100, 100),
    "OperatingCashFlow": _h(300, 330, 380),
    "DistributionsFromAffiliates": _h(50, 55, 70),
    "Capex": _h(-150, -160, -170),
    "InvestingCF": _h(-160, -170, -180),
    "Distributions": _h(-150, -160, -170),
    "DistributionsToNCI": _h(-20, -20, -20),
    "FinancingCF": _h(-120, -150, -180),
    "NetChangeInCash": _h(20, 10, 20),
    "Cash": _h(100, 110, 130),
    "Receivables": _h(100, 110, 120),
    "Inventory": _h(40, 44, 48),
    "NetPPE": _h(2000, 2050, 2100),
    "EquityInvestments": _h(500, 505, 495),
    "TotalAssets": _h(2900, 2950, 3000),
    "Payables": _h(80, 87, 94),
    "TotalDebt": _h(1000, 1010, 1020),
    "LiabilitiesAndEquity": _h(2900, 2950, 3000),
    "EquityCommon": _h(1200, 1220, 1250),
    "NCI": _h(100, 100, 100),
    "TotalEquityInclNCI": _h(1300, 1320, 1350),
}


def _oke_shaped():
    """OKE-like: no TotalCostsAndExpenses (operating cost falls back to REV - OpInc),
    no Liabilities / TotalCurrentLiabilities subtotal -- other liabilities are derived
    from LiabilitiesAndEquity, and include temporary equity not broken out."""
    h = {k: dict(v) for k, v in TOY.items() if k != "TotalCostsAndExpenses"}
    h["TotalEquityInclNCI"] = _h(None, None, None)
    return h


def _close(a, b, tol=1e-6):
    return abs((a or 0.0) - (b or 0.0)) < tol


def _run(da_basis="ppe", history=None, **kw):
    spec = build_spec(da_basis)
    return spec, compute(spec, history or TOY, FY, FCST, **kw)


# --- Review Focus 4 -------------------------------------------------------------------

def test_history_balance_check_zero_with_derived_other_liab():
    _, out = _run(history=_oke_shaped())
    for fy in FY:
        assert _close(out["BAL_CHECK"][fy], 0.0)
        assert _close(out["TOTAL_LE"][fy], TOY["LiabilitiesAndEquity"][fy])
        assert _close(out["TOTAL_ASSETS"][fy], TOY["TotalAssets"][fy])
    # OTHER_LIAB = L&E - AP - DEBT - EQUITY - NCI: 3000 - 94 - 1020 - 1250 - 100
    assert _close(out["OTHER_LIAB"]["FY2025"], 536.0)
    # OPCOST falls back to REV - OpInc - DA - IMPAIR + GAINS: 1200 - 260 - 120 - 0 + 5
    assert _close(out["OPCOST"]["FY2025"], 825.0)
    assert _close(out["OPCOST"]["FY2024"], 1100 - 230 - 110 - 10 + 0)


def test_liabilities_and_equity_falls_back_to_total_assets_when_unmapped():
    h = _oke_shaped()
    del h["LiabilitiesAndEquity"]
    _, out = _run(history=h)
    assert all(_close(out["BAL_CHECK"][fy], 0.0) for fy in FY + FCST)


def test_total_assets_falls_back_to_liabilities_and_equity_when_unmapped():
    h = dict(TOY, TotalAssets=_h(2900, 2950, None))
    _, out = _run(history=h)
    assert _close(out["TOTAL_ASSETS"]["FY2025"], 3000.0)
    assert all(_close(out["BAL_CHECK"][fy], 0.0) for fy in ["FY2025"] + FCST)


def test_history_subtotals_tie_to_reported():
    _, out = _run()
    for fy in FY:
        assert _close(out["PRETAX"][fy], TOY["PretaxIncome"][fy])
        assert _close(out["NI"][fy], TOY["NetIncome"][fy])
        assert _close(out["OCF"][fy], TOY["OperatingCashFlow"][fy])
        assert _close(out["ICF"][fy], TOY["InvestingCF"][fy])
        assert _close(out["FCF_FIN"][fy], TOY["FinancingCF"][fy])
        assert _close(out["NET_CHANGE"][fy], TOY["NetChangeInCash"][fy])
    # TCE path: 940 - 120 - 0 + 5
    assert _close(out["OPCOST"]["FY2025"], 825.0)
    assert _close(out["OTHER_INC"]["FY2025"], 10.0)
    assert _close(out["CAPEX"]["FY2025"], -170.0)


# --- Review Focus 4: first history column ------------------------------------------------

# RATIO lines whose history formula reads a prior column, plus the two CF lines that do.
BLANK_FIRST = {"REV_GROWTH_HIST", "DA_PCT_HIST", "INT_RATE_HIST", "DPU_GROWTH_HIST",
               "DEBT_NET", "CF_DWC"}


def test_blank_first_flag_follows_the_lag_rule():
    assert {ld.key for ld in build_spec("ppe").lines if ld.blank_first} == BLANK_FIRST
    # D&A % revenue reads no prior column, so it keeps its first-year value
    assert {ld.key for ld in build_spec("revenue").lines if ld.blank_first} == \
        BLANK_FIRST - {"DA_PCT_HIST"}


def test_first_history_column_is_blank_on_lag_lines():
    _, out = _run()
    first, rest = FY[0], FY[1:]
    for key in BLANK_FIRST:
        assert out[key][first] is None, key
        assert all(out[key][fy] is not None for fy in rest), key
    assert _close(out["DEBT_NET"]["FY2024"], 1010 - 1000)
    assert _close(out["REV_GROWTH_HIST"]["FY2024"], 0.1)
    # balance check and the derived "other" lines stay numeric
    for key in ("BAL_CHECK", "CF_OTHER_OP", "CF_OTHER_FIN", "OCF", "FCF_FIN"):
        assert out[key][first] is not None, key
    assert _close(out["BAL_CHECK"][first], 0.0)


def test_first_column_other_lines_absorb_so_cash_totals_equal_reported():
    _, out = _run()
    first = FY[0]
    for fy in FY:
        assert _close(out["OCF"][fy], out["OCF_RPT"][fy])
        assert _close(out["FCF_FIN"][fy], out["FCF_RPT"][fy])
    # year 1: the whole working-capital change and debt change sit in the "other" lines
    ops = sum(out[k][first] for k in ("CF_NI", "CF_DA", "CF_IMPAIR", "CF_GAINS", "CF_EQ",
                                      "CF_DIST_AFF"))
    assert _close(out["CF_OTHER_OP"][first], 300 - ops)
    assert _close(out["CF_OTHER_FIN"][first], -120 - (-150) - (-20))


def test_hist_averages_skip_the_blank_first_column():
    d = driver_defaults(build_spec("ppe"), TOY, FY)
    assert _close(d["REV_GROWTH"], (0.1 + 100 / 1100) / 2)
    assert _close(d["DPU_GROWTH"], (0.1 / 1.5 + 0.1 / 1.6) / 2)
    assert _close(d["DA_PCT"], (110 / 2000 + 120 / 2050) / 2)


def test_single_history_year_drivers_read_blank_as_zero():
    one = {k: {"FY2025": v["FY2025"]} for k, v in TOY.items()}
    spec = build_spec("ppe")
    d = driver_defaults(spec, one, ["FY2025"])
    assert d["REV_GROWTH"] == 0.0 and d["INT_RATE"] == 0.0 and d["DPU_GROWTH"] == 0.0
    out = compute(spec, one, ["FY2025"], FCST)
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)


# --- Forecast balance ---------------------------------------------------------------------

def test_forecast_balances_every_year():
    _, out = _run()
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)
    # equity earnings (60) != distributions from affiliates (70): the roll-forward moves EQ_INV
    assert _close(out["EQ_INV"]["FY2026E"], 495 + 60 - 70)


@pytest.mark.parametrize("kw", [
    {"scalar_overrides": {"WC_HOLD": 1}},
    {"scalar_overrides": {"JV_TOGGLE": 1}},
    {"driver_overrides": {"OTHER_OP": 25.0, "DEBT_ISSUE": -40.0}},
    {"driver_overrides": {"REV_GROWTH": {"FY2027E": -0.3}}},
])
def test_forecast_balances_under_overrides(kw):
    _, out = _run(**kw)
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)


@pytest.mark.parametrize("hold", [0, 0.5, 1, 2])
def test_forecast_balances_for_any_wc_hold_value(hold):
    _, out = _run(scalar_overrides={"WC_HOLD": hold})
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)


def test_forecast_balances_oke_shaped_and_revenue_basis():
    _, out = _run(da_basis="revenue", history=_oke_shaped())
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)


def test_wc_hold_keeps_working_capital_flat():
    _, out = _run(scalar_overrides={"WC_HOLD": 1})
    for fy in FCST:
        assert out["AR"][fy] == 120 and out["INV"][fy] == 48 and out["AP"][fy] == 94
        assert out["CF_DWC"][fy] == 0


def test_compute_never_raises_on_missing_history():
    h = {"Revenue": _h(100, None, 120), "Cash": _h(None, None, 10)}
    _, out = _run(history=h)
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)
    assert out["NET_PPE"]["FY2023"] is None          # unmapped concept -> blank
    assert out["CASH_PRE"]["FY2025"] is None          # forecast-only helper -> blank in history
    assert out["DEFTAX"]["FY2026E"] is None           # history-only helper -> blank in forecast


# --- Revolver (Review Focus 5) --------------------------------------------------------------

def test_revolver_draws_shortfall_to_min_cash():
    # DPU growth of 200% makes distributions huge, so cash would fall below MIN_CASH.
    _, out = _run(driver_overrides={"DPU_GROWTH": 2.0})
    min_cash = 130.0
    y1 = FCST[0]
    pre = (out["CASH"]["FY2025"] + out["OCF"][y1] + out["ICF"][y1] + out["DIST"][y1]
           + out["DIST_NCI"][y1] + out["DEBT_NET"][y1] + out["CF_OTHER_FIN"][y1]
           + out["FX_OTHER"][y1])
    assert pre < min_cash
    assert _close(out["CASH_PRE"][y1], pre)
    assert _close(out["REVOLVER"][y1], min_cash - pre)
    for fy in FCST:
        assert _close(out["CASH"][fy], min_cash)
        assert out["REVOLVER_NET"][fy] > 0
        assert abs(out["BAL_CHECK"][fy]) < 1e-6
    # interest is charged on opening revolver at the revolver rate
    y2 = FCST[1]
    d = driver_defaults(build_spec("ppe"), TOY, FY)
    assert _close(out["INTEREST"][y2],
                  d["INT_RATE"] * out["DEBT"][y1] + d["REV_RATE"] * out["REVOLVER"][y1])


def test_revolver_repays_when_cash_exceeds_min():
    # MIN_CASH above last cash: year 1 draws, later years' surplus cash repays the revolver.
    _, out = _run(driver_overrides={"MIN_CASH": 200.0})
    y1, y2 = FCST[0], FCST[1]
    assert out["REVOLVER"][y1] > 0 and _close(out["CASH"][y1], 200.0)
    assert out["REVOLVER_NET"][y2] < 0
    for fy in FCST:
        assert out["REVOLVER"][fy] >= -1e-9               # never repays more than drawn
        assert out["CASH"][fy] >= 200.0 - 1e-6
        if out["REVOLVER"][fy] > 1e-9:
            assert _close(out["CASH"][fy], 200.0)         # surplus goes to repayment first
    assert _close(out["REVOLVER"][FCST[-1]], 0.0)
    assert out["CASH"][FCST[-1]] > 200.0


def test_no_revolver_when_cash_stays_above_min():
    _, out = _run()
    assert all(out["REVOLVER"][fy] == 0 for fy in FCST)
    assert all(out["LIQ_SHORTFALL"][fy] == 0 for fy in FCST)
    assert all(out["LIQ_SHORTFALL"][fy] is None for fy in FY)      # blank in history


def test_revolver_capped_by_commitment_and_shortfall_reported():
    # Huge distributions need more than the 50 commitment: the draw stops at 50, cash
    # falls below minimum by the uncovered amount and the balance sheet still ties.
    _, out = _run(driver_overrides={"DPU_GROWTH": 2.0}, scalar_overrides={"REV_CAP": 50.0})
    _, free = _run(driver_overrides={"DPU_GROWTH": 2.0})
    for fy in FCST:
        assert _close(out["REVOLVER"][fy], 50.0)
        assert out["LIQ_SHORTFALL"][fy] > 0
        assert _close(out["LIQ_SHORTFALL"][fy], 130.0 - out["CASH"][fy])
        assert abs(out["BAL_CHECK"][fy]) < 1e-6
        assert free["REVOLVER"][fy] > 50.0                  # uncapped would draw more
    y1 = FCST[0]
    assert _close(out["REVOLVER_NET"][y1], 50.0)
    assert all(_close(out["REVOLVER_NET"][fy], 0.0) for fy in FCST[1:])


def test_revolver_unlimited_when_cap_zero():
    spec, out = _run(driver_overrides={"DPU_GROWTH": 2.0})
    assert scalar_defaults(spec, TOY, FY)["REV_CAP"] == 0.0
    note = next(x.note for x in spec.scalars if x.key == "REV_CAP")
    assert "forces repayment down to the commitment" in note
    for fy in FCST:
        assert _close(out["LIQ_SHORTFALL"][fy], 0.0)
        assert _close(out["CASH"][fy], 130.0)
        assert abs(out["BAL_CHECK"][fy]) < 1e-6


def _with_revolver():
    """Reported total debt 120/125/130 of which a drawn revolver 10/20/30."""
    base = dict(TOY, TotalDebt=_h(120, 125, 130))
    return base, dict(base, RevolverDrawn=_h(10, 20, 30))


def test_opening_revolver_from_history():
    base_h, h = _with_revolver()
    spec = build_spec("ppe")
    _, base = _run(history=base_h)
    _, out = _run(history=h, driver_overrides={"REV_RATE": 0.08})
    for fy, td, rv in zip(FY, (120, 125, 130), (10, 20, 30)):
        assert _close(out["DEBT"][fy], td - rv)
        assert _close(out["REVOLVER"][fy], rv)
        assert _close(out["DEBT"][fy] + out["REVOLVER"][fy], td)
        assert _close(out["DEBT_NET"][fy], base["DEBT_NET"][fy])       # financing unchanged
        assert _close(out["CF_OTHER_FIN"][fy], base["CF_OTHER_FIN"][fy])
        assert _close(out["FCF_FIN"][fy], TOY["FinancingCF"][fy])
        assert _close(out["OTHER_LIAB"][fy], base["OTHER_LIAB"][fy])
        assert _close(out["BAL_CHECK"][fy], 0.0)
    assert _close(out["DEBT_NET"]["FY2024"], 125 - 120)
    d = driver_defaults(spec, h, FY)
    # interest rate on average (term debt + revolver): unchanged by the split
    assert _close(d["INT_RATE"], 44 / ((130 + 125) / 2))
    assert _close(d["REV_RATE"], d["INT_RATE"])
    y1 = FCST[0]
    assert _close(out["INTEREST"][y1], d["INT_RATE"] * 100 + 0.08 * 30)
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)
    # unmapped revolver -> 0: term debt is the whole reported total
    assert all(_close(base["REVOLVER"][fy], 0.0) for fy in FY)


def test_min_cash_per_year_driver():
    _, out = _run(driver_overrides={"DPU_GROWTH": 2.0,
                                    "MIN_CASH": {"FY2026E": 10.0, "FY2027E": 40.0}})
    assert _close(out["CASH"]["FY2026E"], 10.0)
    for fy in FCST[1:]:                                   # later years chain =prev
        assert _close(out["CASH"][fy], 40.0)
    assert all(_close(out["LIQ_SHORTFALL"][fy], 0.0) for fy in FCST)
    assert all(abs(out["BAL_CHECK"][fy]) < 1e-6 for fy in FCST)


def test_growth_cap_scalar():
    h = dict(TOY, Revenue=_h(100, 200, 400))
    _, out = _run(history=h, scalar_overrides={"GROWTH_CAP": 0.05})
    assert _close(out["REV"]["FY2026E"], 400 * 1.05)
    _, dflt = _run(history=h)
    assert _close(dflt["REV"]["FY2026E"], 400 * 1.15)
    # DPU growth is capped by the same scalar (negative side: -5%)
    h2 = dict(TOY, Distributions=_h(-400, -200, -100))
    _, out2 = _run(history=h2, scalar_overrides={"GROWTH_CAP": 0.05})
    assert _close(out2["DIST"]["FY2026E"], -1.0 * 0.95 * 100)


# --- Toggles and drivers ---------------------------------------------------------------------

def test_jv_toggle_moves_ebitda_not_net_income():
    _, on = _run(scalar_overrides={"JV_TOGGLE": 1})
    _, off = _run()             # default 0: equity earnings excluded
    for fy in FY + FCST:
        assert _close(on["EBITDA"][fy] - off["EBITDA"][fy], on["EQ_EARN"][fy])
        assert _close(on["NI"][fy], off["NI"][fy])
        assert _close(on["EBIT"][fy], off["EBIT"][fy])
    assert on["EQ_EARN"]["FY2026E"] == 60


def test_da_basis_revenue_when_ppe_missing():
    h = {k: v for k, v in TOY.items() if k != "NetPPE"}
    spec, out = _run(da_basis="revenue", history=h)
    d = driver_defaults(spec, h, FY)
    assert _close(d["DA_PCT"], (100 / 1000 + 110 / 1100 + 120 / 1200) / 3)
    for fy in FCST:
        assert _close(out["DA"][fy], d["DA_PCT"] * out["REV"][fy])
        assert _close(out["DA_PCT_HIST"][fy], d["DA_PCT"])
    # ppe basis uses opening net PP&E instead
    spec_p, out_p = _run()
    dp = driver_defaults(spec_p, TOY, FY)
    assert _close(out_p["DA"]["FY2026E"], dp["DA_PCT"] * 2100)


def test_nci_flat_by_default():
    _, out = _run()
    for fy in FCST:
        assert _close(out["NCI_BS"][fy], 100.0)
        assert _close(out["DIST_NCI"][fy], -out["NI_NCI"][fy])
        assert _close(out["NI_NCI"][fy], 20.0)


def test_tax_rate_default_is_effective_rate_even_with_deferred_tax():
    """Spec 7.2: the TAX_RATE default is the effective rate tax / pretax (3-yr average,
    clamped 0-35%). Deferred tax does not lower it: deferred taxes reverse."""
    spec = build_spec("ppe")
    effective = (10 / 210 + 12 / 238 + 14 / 286) / 3
    assert _close(driver_defaults(spec, TOY, FY)["TAX_RATE"], effective)
    h = dict(TOY, DeferredTax=_h(4, 5, 6))
    with_def = driver_defaults(spec, h, FY)["TAX_RATE"]
    assert _close(with_def, effective)
    out = compute(spec, h, FY, FCST)
    for fy in FCST:
        assert _close(out["TAX"][fy], with_def * out["PRETAX"][fy])
    # clamped to [0, 35%]
    h2 = dict(TOY, IncomeTax=_h(200, 200, 200), NetIncome=_h(10, 38, 86))
    assert _close(driver_defaults(spec, h2, FY)["TAX_RATE"], 0.35)
    h3 = dict(TOY, IncomeTax=_h(-50, -50, -50), NetIncome=_h(260, 288, 336))
    assert driver_defaults(spec, h3, FY)["TAX_RATE"] == 0.0


def test_unmapped_pretax_income_falls_back_to_net_income_plus_tax():
    h = dict(TOY, PretaxIncome=_h(None, None, None))
    _, base = _run()
    _, out = _run(history=h)
    for fy in FY:
        assert _close(out["PRETAX"][fy], TOY["NetIncome"][fy] + TOY["IncomeTax"][fy])
    for fy in FCST:
        assert abs(out["NI"][fy] - base["NI"][fy]) <= 0.01 * abs(base["NI"][fy])
        assert abs(out["BAL_CHECK"][fy]) < 1e-6


def test_cash_tax_rate_info_line_differs_when_deferred_tax_present():
    spec = build_spec("ppe")
    ld = spec.line("CASH_TAX_RATE_HIST")
    assert ld.section == "RATIO" and not ld.hidden and ld.label == "Cash tax rate (info)"
    assert "effective rate" in ld.note
    out = compute(spec, TOY, FY, FCST)                # no deferred tax: the two agree
    for fy in FY:
        assert _close(out["CASH_TAX_RATE_HIST"][fy], out["TAX_RATE_HIST"][fy])
    h = dict(TOY, DeferredTax=_h(4, 5, 6))
    out = compute(spec, h, FY, FCST)
    for fy, tax, d, pre in zip(FY, (10, 12, 14), (4, 5, 6), (210, 238, 286)):
        assert _close(out["CASH_TAX_RATE_HIST"][fy], (tax - d) / pre)
        assert _close(out["TAX_RATE_HIST"][fy], tax / pre)
        assert not _close(out["CASH_TAX_RATE_HIST"][fy], out["TAX_RATE_HIST"][fy])


def test_driver_defaults_from_history():
    d = driver_defaults(build_spec("ppe"), TOY, FY)
    assert set(d) == {"REV_GROWTH", "COST_PCT", "DA_PCT", "CAPEX_PCT", "INT_RATE", "TAX_RATE",
                      "DPU_GROWTH", "UNITS", "AR_DAYS", "INV_DAYS", "AP_DAYS", "OTHER_OP",
                      "DEBT_ISSUE", "MIN_CASH", "REV_RATE"}
    assert [x.key for x in build_spec("ppe").drivers][-3:] == ["DEBT_ISSUE", "MIN_CASH", "REV_RATE"]
    assert d["MIN_CASH"] == 130.0 and _close(d["REV_RATE"], d["INT_RATE"])
    # the first history column is blank (no prior year), so the average skips it
    assert _close(d["REV_GROWTH"], (0.1 + 100 / 1100) / 2)
    assert _close(d["COST_PCT"], 825 / 1200)
    assert _close(d["CAPEX_PCT"], (150 / 1000 + 160 / 1100 + 170 / 1200) / 3)
    assert _close(d["INT_RATE"], 44 / 1015)
    assert _close(d["DPU_GROWTH"], (0.1 / 1.5 + 0.1 / 1.6) / 2)
    assert d["UNITS"] == 100 and d["OTHER_OP"] == 0 and d["DEBT_ISSUE"] == 0
    assert _close(d["AR_DAYS"], 120 / 1200 * 365)
    assert _close(d["INV_DAYS"], 48 / 825 * 365)
    assert _close(d["AP_DAYS"], 94 / 825 * 365)


def test_growth_drivers_capped_at_15_percent():
    h = dict(TOY, Revenue=_h(100, 200, 400))
    assert driver_defaults(build_spec("ppe"), h, FY)["REV_GROWTH"] == 0.15


def test_scalar_defaults():
    s = scalar_defaults(build_spec("ppe"), TOY, FY)
    assert s == {"JV_TOGGLE": 0.0, "WC_HOLD": 0.0, "REV_CAP": 0.0, "GROWTH_CAP": 0.15}


def test_default_asset_block_and_distributions():
    spec, out = _run()
    d = driver_defaults(spec, TOY, FY)
    for n, fy in enumerate(FCST, start=1):
        rev = 1200 * (1 + d["REV_GROWTH"]) ** n
        assert _close(out["REV"][fy], rev)
        assert _close(out["OPCOST"][fy], d["COST_PCT"] * rev)
        dpu = 1.7 * (1 + d["DPU_GROWTH"]) ** n
        assert _close(out["DIST"][fy], -dpu * 100)
        assert _close(out["CAPEX"][fy], -d["CAPEX_PCT"] * rev)


def test_per_year_driver_override_chains_forward():
    _, out = _run(driver_overrides={"REV_GROWTH": {"FY2027E": 0.0}})
    assert _close(out["REV"]["FY2027E"], out["REV"]["FY2026E"])
    assert _close(out["REV"]["FY2030E"], out["REV"]["FY2026E"])   # later years chain =prev


def test_explicit_asset_block():
    blk = {"ASSET_REV": {fy: 1500.0 for fy in FCST}, "ASSET_COST": {fy: 900.0 for fy in FCST}}
    _, out = _run(asset_block=blk)
    assert out["REV"]["FY2028E"] == 1500 and out["ASSET_EBITDA"]["FY2028E"] == 600


def test_unknown_override_key_raises():
    with pytest.raises(KeyError):
        _run(driver_overrides={"NOPE": 1.0})


# --- Spec structure ----------------------------------------------------------------------------

IS_KEYS = ["REV", "OPCOST", "ASSET_EBITDA", "EQ_EARN", "EBITDA", "DA", "IMPAIR", "GAINS", "EBIT",
           "INTEREST", "OTHER_INC", "PRETAX", "TAX", "DISC_OPS", "NI", "NI_NCI", "NI_COMMON",
           "UNITS", "EPS"]
CF_KEYS = ["CF_NI", "CF_DA", "CF_IMPAIR", "CF_GAINS", "CF_EQ", "CF_DIST_AFF", "CF_DWC",
           "CF_OTHER_OP", "OCF", "CAPEX", "CF_OTHER_INV", "ICF", "DIST", "DIST_NCI", "DEBT_NET",
           "REVOLVER_NET", "CF_OTHER_FIN", "FCF_FIN", "FX_OTHER", "NET_CHANGE"]
BS_KEYS = ["CASH", "AR", "INV", "NET_PPE", "EQ_INV", "OTHER_ASSETS", "TOTAL_ASSETS", "AP", "DEBT",
           "REVOLVER", "OTHER_LIAB", "EQUITY", "NCI_BS", "TOTAL_LE"]
RATIO_KEYS = ["REV_GROWTH_HIST", "EBITDA_MARGIN", "COST_PCT_HIST", "DA_PCT_HIST",
              "CAPEX_PCT_HIST", "INT_RATE_HIST", "TAX_RATE_HIST", "CASH_TAX_RATE_HIST", "DPU_HIST",
              "DPU_GROWTH_HIST",
              "AR_DAYS_HIST", "INV_DAYS_HIST", "AP_DAYS_HIST", "NET_DEBT_EBITDA", "DIST_COVERAGE"]


def test_visible_lines_in_display_order():
    spec = build_spec("ppe")
    by_section = {}
    for ld in spec.lines:
        if not ld.hidden:
            by_section.setdefault(ld.section, []).append(ld.key)
    assert by_section == {"IS": IS_KEYS, "CF": CF_KEYS, "BS": BS_KEYS,
                          "CHECK": ["BAL_CHECK", "LIQ_SHORTFALL"], "RATIO": RATIO_KEYS}
    order = [ld.section for ld in spec.lines]
    assert order == sorted(order, key=["IS", "CF", "BS", "CHECK", "RATIO"].index)
    assert {"CASH_PRE", "DEFTAX", "DPU_FCST", "REV_RPT"} <=         {ld.key for ld in spec.lines if ld.hidden}
    assert spec.line("REV_RPT").hist == "RevolverDrawn"
    assert spec.line("LIQ_SHORTFALL").label == "Liquidity shortfall vs minimum cash"


def test_ratio_lines_compute_in_both_ranges():
    for ld in build_spec("ppe").lines:
        if ld.section == "RATIO":
            assert ld.hist is not None and ld.hist == ld.fcst


@pytest.mark.parametrize("basis", ["ppe", "revenue"])
def test_same_column_dependency_graph_is_acyclic(basis):
    spec = build_spec(basis)
    validate_spec(spec)
    keys = spec.line_keys()
    for phase in ("hist", "fcst"):
        order = eval_order(spec, phase)
        assert sorted(order) == sorted(keys)
        pos = {k: i for i, k in enumerate(order)}
        for ld in spec.lines:
            f = ld.hist if phase == "hist" else ld.fcst
            if f is None or isinstance(f, str):
                continue
            for node in iter_nodes(f):
                if isinstance(node, Line) and node.lag == 0:
                    assert pos[node.key] < pos[ld.key], (phase, ld.key, node.key)


def test_cycle_is_rejected():
    spec = build_spec("ppe")
    bad = ModelSpec(lines=spec.lines + [
        LineDef("A", "a", "IS", None, Line("B")), LineDef("B", "b", "IS", None, Line("A"))],
        drivers=spec.drivers, scalars=spec.scalars, da_basis="ppe")
    with pytest.raises(ValueError, match="circular"):
        eval_order(bad, "fcst")


def test_validate_rejects_unknown_reference():
    spec = build_spec("ppe")
    bad = ModelSpec(lines=spec.lines + [LineDef("Z", "z", "IS", None, Line("NOPE", 1))],
                    drivers=spec.drivers, scalars=spec.scalars, da_basis="ppe")
    with pytest.raises(ValueError, match="NOPE"):
        validate_spec(bad)


def test_every_formula_renders_to_excel():
    ctx = RenderContext(col=5, line_cell=lambda k, c: f"R{c}{k}", driver_cell=lambda k, c: k,
                        scalar_cell=lambda k: k, block_range=lambda k, c: k,
                        hist_range=lambda k, n: k)
    spec = build_spec("ppe")
    exprs = [f for ld in spec.lines for f in (ld.hist, ld.fcst) if f is not None
             and not isinstance(f, str)]
    exprs += [d.default for d in spec.drivers]
    for e in exprs:
        assert to_excel(e, ctx).startswith("=")
        for node in iter_nodes(e):
            if isinstance(node, Fn) and node.name == "IFERROR":
                assert len(node.args) == 2


def test_spec_rejects_bad_basis():
    with pytest.raises(ValueError):
        build_spec("assets")


def test_outputs_are_finite():
    _, out = _run()
    for key, row in out.items():
        for fy, v in row.items():
            assert v is None or math.isfinite(v), (key, fy)
