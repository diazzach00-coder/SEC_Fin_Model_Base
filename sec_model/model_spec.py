"""The linked 3-statement model, written once as expressions over named lines.

Each `LineDef` carries a history formula and a forecast formula:
- `hist` is a concept name (history cells pull that concept from the Mapping), an
  `Expr` (derived from other lines in the same history column), or None (blank).
- `fcst` is an `Expr` evaluated in each forecast column, or None (blank).

`MODEL_LINES` is in display order (IS, CF, BS, CHECK, RATIO). Evaluation order is
NOT list order: within one column, lag-0 `Line` references form a dependency graph
that must be acyclic (asserted by tests), and `forecast.compute` evaluates each
column in topological order. Lag-1 references always point at the previous,
already-computed column. Hidden helper lines (`hidden=True`) are real rows the
workbook hides: `<X>_RPT` rows pull a reported concept that a derived line needs,
`DEFTAX` feeds the informational cash-tax-rate ratio, `CASH_PRE` feeds the revolver and `DPU_FCST`
carries distributions per unit forward.

Debt is split into term debt (`DEBT` = reported total debt less the drawn revolver) and
the revolver (`REVOLVER`, opening balance = the reported drawn revolver, 0 when untagged).
The revolver draws to hold the per-year `MIN_CASH` driver, up to the `REV_CAP` commitment
(0 = unlimited); `LIQ_SHORTFALL` reports cash left below the minimum.

Sign convention: IS expenses positive on lines named as expenses (OPCOST, DA,
IMPAIR, INTEREST, TAX); GAINS positive = gain; CF lines carry their cash sign.

Drivers fill the first forecast column of the Drivers sheet with their `default`
(evaluated in the last history column, so `Line` means last year and `Hist`
averages the last n history years of a RATIO line); later columns chain `=prev`.
"""
from __future__ import annotations

from dataclasses import dataclass

from .formula import Block, Const, Driver, Expr, Fn, Hist, Line, Scalar

SECTIONS = ("IS", "CF", "BS", "CHECK", "RATIO")
FMTS = ("money", "pct", "units", "perunit", "days", "multiple")
DA_BASES = ("ppe", "revenue")
BLOCKS = ("ASSET_REV", "ASSET_COST")


@dataclass
class LineDef:
    key: str
    label: str
    section: str
    hist: str | Expr | None
    fcst: Expr | None
    fmt: str = "money"
    bold: bool = False
    is_input_row: bool = False
    hidden: bool = False
    note: str = ""


@dataclass
class DriverDef:
    key: str
    label: str
    default: Expr
    fmt: str = "pct"
    note: str = ""


@dataclass
class ScalarDef:
    key: str
    label: str
    default: float | Expr
    fmt: str = "units"
    note: str = ""


@dataclass
class ModelSpec:
    lines: list[LineDef]
    drivers: list[DriverDef]
    scalars: list[ScalarDef]
    da_basis: str

    def line(self, key: str) -> LineDef:
        for ld in self.lines:
            if ld.key == key:
                return ld
        raise KeyError(key)

    def line_keys(self) -> list[str]:
        return [ld.key for ld in self.lines]


def L(key: str, lag: int = 0) -> Line:
    return Line(key, lag)


def _sum(*keys: str) -> Expr:
    e = L(keys[0])
    for k in keys[1:]:
        e = e + L(k)
    return e


def _neg_abs(key: str) -> Expr:
    return -Fn("ABS", L(key))


def _growth(key: str) -> Expr:
    """x/x_prev - 1, but 0 when there is no prior value (first history column)."""
    return Fn("IF", L(key, 1).eq(0), 0, L(key) / L(key, 1) - 1)


def _capped(e: Expr) -> Expr:
    """Clamp a default growth rate to +/- the GROWTH_CAP scalar."""
    cap = Scalar("GROWTH_CAP")
    return Fn("MAX", Fn("MIN", e, cap), -cap)


def _wc(key: str, days: str, base: str) -> Expr:
    """Working-capital balance off days; held at prior when WC_HOLD = 1."""
    return Fn("IF", Scalar("WC_HOLD").eq(1), L(key, 1), Driver(days) / 365 * L(base))


def _rpt(key: str, concept: str, section: str) -> LineDef:
    return LineDef(key, f"{concept} (as reported)", section, concept, None, hidden=True,
                   note="hidden helper: reported concept used by a derived history line")


def _same(key, label, expr, fmt="money", bold=False, section="IS", note="") -> LineDef:
    return LineDef(key, label, section, expr, expr, fmt=fmt, bold=bold, note=note)


def _ratio(key, label, expr, fmt="pct", note="") -> LineDef:
    return LineDef(key, label, "RATIO", expr, expr, fmt=fmt, note=note)


def _is_lines(da_basis: str) -> list[LineDef]:
    jv = Scalar("JV_TOGGLE")
    da_fcst = Driver("DA_PCT") * (L("NET_PPE", 1) if da_basis == "ppe" else L("REV"))
    opcost_hist = (Fn("IF", L("TCE_RPT").eq(0), L("REV") - L("OPINC_RPT"), L("TCE_RPT"))
                   - L("DA") - L("IMPAIR") + L("GAINS"))
    # Unmapped pretax income (blank reads 0) falls back to net income + tax; otherwise
    # OTHER_INC absorbs -EBIT + interest and the held-flat forecast wipes out pretax.
    pretax_hist = Fn("IF", L("PRETAX_RPT").eq(0), L("NI_RPT") + L("TAX"), L("PRETAX_RPT"))
    return [
        LineDef("REV", "Revenue", "IS", "Revenue", Block("ASSET_REV"), bold=True),
        _rpt("TCE_RPT", "TotalCostsAndExpenses", "IS"),
        _rpt("OPINC_RPT", "OperatingIncome", "IS"),
        LineDef("OPCOST", "Operating cost (excl. D&A, impairments, gains)", "IS", opcost_hist,
                Block("ASSET_COST"),
                note="total costs - D&A - impairments + gains; falls back to "
                     "revenue - operating income when total costs are not reported"),
        _same("ASSET_EBITDA", "Asset EBITDA", L("REV") - L("OPCOST"), bold=True),
        LineDef("EQ_EARN", "Equity earnings from unconsolidated affiliates", "IS",
                "EquityEarnings", L("EQ_EARN", 1), note="held flat"),
        _same("EBITDA", "EBITDA", L("ASSET_EBITDA") + L("EQ_EARN") * jv, bold=True,
              note="includes equity earnings when JV_TOGGLE = 1"),
        LineDef("DA", "Depreciation & amortization", "IS", "DA", da_fcst,
                note="% of opening net PP&E" if da_basis == "ppe" else "% of revenue"),
        LineDef("IMPAIR", "Impairments", "IS", "Impairments", Const(0)),
        LineDef("GAINS", "Gains (losses) on asset sales", "IS", "Gains", Const(0)),
        _same("EBIT", "EBIT", L("EBITDA") + L("EQ_EARN") * (1 - jv) - L("DA") - L("IMPAIR")
              + L("GAINS"), bold=True),
        LineDef("INTEREST", "Interest expense", "IS", "InterestExpense",
                Driver("INT_RATE") * L("DEBT", 1) + Driver("REV_RATE") * L("REVOLVER", 1),
                note="term-debt rate on opening term debt + revolver rate on opening revolver "
                     "(no circularity)"),
        _rpt("PRETAX_RPT", "PretaxIncome", "IS"),
        LineDef("OTHER_INC", "Other income (expense)", "IS",
                pretax_hist - L("EBIT") + L("INTEREST"), L("OTHER_INC", 1),
                note="derived in history (pretax falls back to net income + tax when "
                     "unreported); held flat"),
        _same("PRETAX", "Pretax income", L("EBIT") - L("INTEREST") + L("OTHER_INC"), bold=True),
        LineDef("TAX", "Income tax", "IS", "IncomeTax", Driver("TAX_RATE") * L("PRETAX")),
        LineDef("DEFTAX", "Deferred income tax (CF)", "IS", "DeferredTax", None, hidden=True,
                note="hidden helper for CASH_TAX_RATE_HIST; blank (0) when unmapped"),
        _rpt("NI_RPT", "NetIncome", "IS"),
        LineDef("DISC_OPS", "Discontinued ops & other", "IS",
                L("NI_RPT") - (L("PRETAX") - L("TAX")), Const(0)),
        _same("NI", "Net income", L("PRETAX") - L("TAX") + L("DISC_OPS"), bold=True),
        LineDef("NI_NCI", "NI to noncontrolling interests", "IS", "NIToNCI", L("NI_NCI", 1),
                note="held flat"),
        _same("NI_COMMON", "NI to common / partners", L("NI") - L("NI_NCI"), bold=True),
        LineDef("UNITS", "Diluted units", "IS", "DilutedUnits", Driver("UNITS"), fmt="units"),
        _same("EPS", "EPS / EPU", L("NI_COMMON") / L("UNITS"), fmt="perunit"),
    ]


def _cf_lines() -> list[LineDef]:
    dwc = -((L("AR") - L("AR", 1)) + (L("INV") - L("INV", 1)) - (L("AP") - L("AP", 1)))
    op_items = ("CF_NI", "CF_DA", "CF_IMPAIR", "CF_GAINS", "CF_EQ", "CF_DIST_AFF", "CF_DWC")
    need = Fn("MAX", Driver("MIN_CASH") - L("CASH_PRE"), -L("REVOLVER", 1))
    revolver_net = Fn("IF", Scalar("REV_CAP").eq(0), need,
                      Fn("MIN", need, Scalar("REV_CAP") - L("REVOLVER", 1)))
    cash_pre = (L("CASH", 1) + L("OCF") + L("ICF") + L("DIST") + L("DIST_NCI") + L("DEBT_NET")
                + L("CF_OTHER_FIN") + L("FX_OTHER"))
    return [
        _same("CF_NI", "Net income", L("NI"), section="CF"),
        _same("CF_DA", "D&A", L("DA"), section="CF"),
        _same("CF_IMPAIR", "Impairments", L("IMPAIR"), section="CF"),
        _same("CF_GAINS", "(Gains) losses on asset sales", -L("GAINS"), section="CF"),
        _same("CF_EQ", "Less equity earnings", -L("EQ_EARN"), section="CF"),
        LineDef("CF_DIST_AFF", "Distributions from affiliates", "CF",
                "DistributionsFromAffiliates", L("CF_DIST_AFF", 1), note="held flat"),
        LineDef("CF_DWC", "Change in working capital", "CF", dwc, dwc,
                note="zero when WC_HOLD = 1 because AR/INV/AP hold at prior"),
        _rpt("OCF_RPT", "OperatingCashFlow", "CF"),
        LineDef("CF_OTHER_OP", "Other operating items", "CF",
                L("OCF_RPT") - _sum(*op_items), Driver("OTHER_OP"),
                note="derived in history; accrues to other liabilities in forecast"),
        _same("OCF", "Cash from operations", _sum(*op_items, "CF_OTHER_OP"), bold=True,
              section="CF"),
        _rpt("CAPEX_RPT", "Capex", "CF"),
        LineDef("CAPEX", "Capital expenditures", "CF", _neg_abs("CAPEX_RPT"),
                -Driver("CAPEX_PCT") * L("REV")),
        _rpt("ICF_RPT", "InvestingCF", "CF"),
        LineDef("CF_OTHER_INV", "Other investing", "CF", L("ICF_RPT") - L("CAPEX"), Const(0)),
        _same("ICF", "Cash from investing", L("CAPEX") + L("CF_OTHER_INV"), bold=True,
              section="CF"),
        _rpt("DIST_RPT", "Distributions", "CF"),
        LineDef("DPU_FCST", "Distribution per unit", "CF",
                Fn("ABS", L("DIST_RPT")) / L("UNITS"),
                L("DPU_FCST", 1) * (1 + Driver("DPU_GROWTH")), fmt="perunit", hidden=True,
                note="hidden helper: last DPU grown by DPU_GROWTH each forecast year"),
        LineDef("DIST", "Distributions / dividends paid", "CF", _neg_abs("DIST_RPT"),
                -L("DPU_FCST") * L("UNITS")),
        _rpt("DIST_NCI_RPT", "DistributionsToNCI", "CF"),
        LineDef("DIST_NCI", "Distributions to NCI", "CF", _neg_abs("DIST_NCI_RPT"),
                -L("NI_NCI"), note="= -NI to NCI so NCI stays flat"),
        LineDef("DEBT_NET", "Debt issuance (repayment)", "CF",
                (L("DEBT") + L("REVOLVER")) - (L("DEBT", 1) + L("REVOLVER", 1)),
                Driver("DEBT_ISSUE"),
                note="history: change in total debt incl. revolver; forecast: term debt only"),
        LineDef("CASH_PRE", "Cash before revolver", "CF", None, cash_pre, hidden=True,
                note="hidden helper: opening cash + every cash line except the revolver"),
        LineDef("REVOLVER_NET", "Revolver draw (repayment)", "CF", Const(0), revolver_net,
                note="draws to hold minimum cash, up to the REV_CAP commitment (0 = unlimited); "
                     "repays when cash exceeds it"),
        _rpt("FCF_RPT", "FinancingCF", "CF"),
        LineDef("CF_OTHER_FIN", "Other financing", "CF",
                L("FCF_RPT") - L("DIST") - L("DIST_NCI") - L("DEBT_NET"), Const(0)),
        _same("FCF_FIN", "Cash from financing",
              _sum("DIST", "DIST_NCI", "DEBT_NET", "REVOLVER_NET", "CF_OTHER_FIN"), bold=True,
              section="CF"),
        _rpt("NETCHG_RPT", "NetChangeInCash", "CF"),
        LineDef("FX_OTHER", "FX & other", "CF",
                L("NETCHG_RPT") - L("OCF") - L("ICF") - L("FCF_RPT"), Const(0)),
        _same("NET_CHANGE", "Net change in cash", _sum("OCF", "ICF", "FCF_FIN", "FX_OTHER"),
              bold=True, section="CF"),
    ]


def _bs_lines() -> list[LineDef]:
    # Each balance-sheet total falls back to the other when unmapped, so the
    # history BAL_CHECK is 0 by construction whenever either total is reported.
    ta = Fn("IF", L("TA_RPT").eq(0), L("LE_RPT"), L("TA_RPT"))
    le = Fn("IF", L("LE_RPT").eq(0), L("TA_RPT"), L("LE_RPT"))
    return [
        LineDef("CASH", "Cash", "BS", "Cash", L("CASH", 1) + L("NET_CHANGE")),
        LineDef("AR", "Receivables", "BS", "Receivables", _wc("AR", "AR_DAYS", "REV")),
        LineDef("INV", "Inventory", "BS", "Inventory", _wc("INV", "INV_DAYS", "OPCOST")),
        LineDef("NET_PPE", "Net PP&E", "BS", "NetPPE",
                L("NET_PPE", 1) - L("CAPEX") - L("DA")),
        LineDef("EQ_INV", "Equity investments", "BS", "EquityInvestments",
                L("EQ_INV", 1) + L("EQ_EARN") - L("CF_DIST_AFF"),
                note="equity-method roll-forward: + equity earnings - distributions received"),
        _rpt("TA_RPT", "TotalAssets", "BS"),
        LineDef("OTHER_ASSETS", "Other assets", "BS",
                ta - _sum("CASH", "AR", "INV", "NET_PPE", "EQ_INV"),
                L("OTHER_ASSETS", 1), note="derived in history; held flat"),
        _same("TOTAL_ASSETS", "Total assets",
              _sum("CASH", "AR", "INV", "NET_PPE", "EQ_INV", "OTHER_ASSETS"), bold=True,
              section="BS"),
        LineDef("AP", "Payables", "BS", "Payables", _wc("AP", "AP_DAYS", "OPCOST")),
        _rpt("DEBT_RPT", "TotalDebt", "BS"),
        _rpt("REV_RPT", "RevolverDrawn", "BS"),
        LineDef("DEBT", "Term debt (excl. revolver)", "BS", L("DEBT_RPT") - L("REV_RPT"),
                L("DEBT", 1) + L("DEBT_NET"),
                note="reported total debt less the drawn revolver"),
        LineDef("REVOLVER", "Revolver", "BS", L("REV_RPT"), L("REVOLVER", 1) + L("REVOLVER_NET"),
                note="opening balance = reported drawn revolver (0 when not tagged)"),
        _rpt("LE_RPT", "LiabilitiesAndEquity", "BS"),
        LineDef("OTHER_LIAB", "Other liabilities", "BS",
                le - L("AP") - L("DEBT") - L("REVOLVER") - L("EQUITY") - L("NCI_BS"),
                L("OTHER_LIAB", 1) + L("CF_OTHER_OP"),
                note="includes preferred/temporary equity and anything not broken out; "
                     "other operating cash items accrue here so the OTHER_OP driver cannot "
                     "unbalance the sheet"),
        LineDef("EQUITY", "Common / partners' equity", "BS", "EquityCommon",
                L("EQUITY", 1) + L("NI_COMMON") + L("DIST")),
        LineDef("NCI_BS", "Noncontrolling interests", "BS", "NCI",
                L("NCI_BS", 1) + L("NI_NCI") + L("DIST_NCI")),
        _same("TOTAL_LE", "Total liabilities & equity",
              _sum("AP", "DEBT", "REVOLVER", "OTHER_LIAB", "EQUITY", "NCI_BS"), bold=True,
              section="BS"),
        _same("BAL_CHECK", "Balance check (assets - L&E)",
              L("TOTAL_ASSETS") - L("TOTAL_LE"), section="CHECK"),
        LineDef("LIQ_SHORTFALL", "Liquidity shortfall vs minimum cash", "CHECK", None,
                Fn("MAX", 0, Driver("MIN_CASH") - L("CASH")),
                note="cash below minimum: revolver commitment exhausted"),
    ]


def _ratio_lines(da_basis: str) -> list[LineDef]:
    da_base = L("NET_PPE", 1) if da_basis == "ppe" else L("REV")
    return [
        _ratio("REV_GROWTH_HIST", "Revenue growth", _growth("REV")),
        _ratio("EBITDA_MARGIN", "EBITDA margin", L("EBITDA") / L("REV")),
        _ratio("COST_PCT_HIST", "Operating cost % revenue", L("OPCOST") / L("REV")),
        _ratio("DA_PCT_HIST",
               "D&A % opening net PP&E" if da_basis == "ppe" else "D&A % revenue",
               L("DA") / da_base),
        _ratio("CAPEX_PCT_HIST", "Capex % revenue", -L("CAPEX") / L("REV")),
        _ratio("INT_RATE_HIST", "Interest rate on average debt",
               L("INTEREST") / Fn("AVERAGE", L("DEBT") + L("REVOLVER"),
                                  L("DEBT", 1) + L("REVOLVER", 1))),
        # The TAX_RATE driver uses the effective rate (spec 7.2): deferred taxes reverse,
        # so a cash rate (OKE ~1.5%) understates long-run tax in the DCF.
        _ratio("TAX_RATE_HIST", "Effective tax rate",
               Fn("MAX", 0, Fn("MIN", 0.35, L("TAX") / L("PRETAX")))),
        _ratio("CASH_TAX_RATE_HIST", "Cash tax rate (info)",
               Fn("MAX", 0, Fn("MIN", 0.35, (L("TAX") - L("DEFTAX")) / L("PRETAX"))),
               note="deferred taxes reverse; driver uses the effective rate - override on "
                    "Drivers if you want cash taxes"),
        _ratio("DPU_HIST", "Distribution per unit", -L("DIST") / L("UNITS"), fmt="perunit"),
        _ratio("DPU_GROWTH_HIST", "DPU growth", _growth("DPU_HIST")),
        _ratio("AR_DAYS_HIST", "Receivable days", L("AR") / L("REV") * 365, fmt="days"),
        _ratio("INV_DAYS_HIST", "Inventory days", L("INV") / L("OPCOST") * 365, fmt="days"),
        _ratio("AP_DAYS_HIST", "Payable days", L("AP") / L("OPCOST") * 365, fmt="days"),
        _ratio("NET_DEBT_EBITDA", "Net debt / EBITDA",
               (L("DEBT") + L("REVOLVER") - L("CASH")) / L("EBITDA"), fmt="multiple"),
        _ratio("DIST_COVERAGE", "Distribution coverage (OCF - capex) / distributions",
               (L("OCF") + L("CAPEX")) / -L("DIST"), fmt="multiple"),
    ]


def _drivers(da_basis: str) -> list[DriverDef]:
    return [
        DriverDef("REV_GROWTH", "Revenue growth (placeholder asset)",
                  _capped(Hist("REV_GROWTH_HIST", 3)),
                  note="3-yr average growth, capped at +/- GROWTH_CAP; drives the placeholder "
                       "asset only"),
        DriverDef("COST_PCT", "Operating cost % revenue (placeholder asset)",
                  L("COST_PCT_HIST"), note="last year"),
        DriverDef("DA_PCT", "D&A % opening net PP&E" if da_basis == "ppe" else "D&A % revenue",
                  Hist("DA_PCT_HIST", 3), note="3-yr average"),
        DriverDef("CAPEX_PCT", "Capex % revenue", Hist("CAPEX_PCT_HIST", 3),
                  note="3-yr average"),
        DriverDef("INT_RATE", "Interest rate on opening term debt", L("INT_RATE_HIST"),
                  note="last year interest / average debt (incl. revolver)"),
        DriverDef("TAX_RATE", "Effective tax rate", Hist("TAX_RATE_HIST", 3),
                  note="3-yr average of tax / pretax, clamped 0-35% (cash rate shown on "
                       "the Model for reference)"),
        DriverDef("DPU_GROWTH", "Distribution per unit growth",
                  _capped(Hist("DPU_GROWTH_HIST", 3)),
                  note="3-yr average growth, capped at +/- GROWTH_CAP"),
        DriverDef("UNITS", "Diluted units", L("UNITS"), fmt="units", note="last year, flat"),
        DriverDef("AR_DAYS", "Receivable days (of revenue)", L("AR_DAYS_HIST"), fmt="days",
                  note="last year"),
        DriverDef("INV_DAYS", "Inventory days (of operating cost)", L("INV_DAYS_HIST"),
                  fmt="days", note="last year"),
        DriverDef("AP_DAYS", "Payable days (of operating cost)", L("AP_DAYS_HIST"), fmt="days",
                  note="last year"),
        DriverDef("OTHER_OP", "Other operating cash items", Const(0), fmt="money",
                  note="input; accrues to other liabilities"),
        DriverDef("DEBT_ISSUE", "Debt issuance (repayment)", Const(0), fmt="money",
                  note="input"),
        DriverDef("MIN_CASH", "Minimum cash balance", L("CASH"), fmt="money",
                  note="last reported cash; the revolver draws/repays to hold it"),
        DriverDef("REV_RATE", "Revolver interest rate", L("INT_RATE_HIST"),
                  note="revolver rate; defaults to the term-debt rate"),
    ]


def _scalars() -> list[ScalarDef]:
    return [
        ScalarDef("JV_TOGGLE", "Include JV equity earnings in EBITDA", 1.0,
                  note="1 = include JV equity earnings in EBITDA, 0 = exclude"),
        ScalarDef("WC_HOLD", "Hold working capital flat", 0.0,
                  note="1 = hold working capital flat (no change in forecast)"),
        ScalarDef("REV_CAP", "Revolver commitment", 0.0, fmt="money",
                  note="revolver commitment; 0 = unlimited"),
        ScalarDef("GROWTH_CAP", "Cap on default growth rates", 0.15, fmt="pct",
                  note="cap on default growth rates"),
    ]


def build_spec(da_basis: str) -> ModelSpec:
    """The full model for one D&A basis: "ppe" (% of opening net PP&E) or "revenue"."""
    if da_basis not in DA_BASES:
        raise ValueError(f"da_basis must be one of {DA_BASES}, got {da_basis!r}")
    lines = _is_lines(da_basis) + _cf_lines() + _bs_lines() + _ratio_lines(da_basis)
    return ModelSpec(lines=lines, drivers=_drivers(da_basis), scalars=_scalars(),
                     da_basis=da_basis)


_DEFAULT = build_spec("ppe")
MODEL_LINES: list[LineDef] = _DEFAULT.lines
DRIVERS: list[DriverDef] = _DEFAULT.drivers
SCALARS: list[ScalarDef] = _DEFAULT.scalars
