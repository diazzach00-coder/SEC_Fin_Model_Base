"""Streamlit UI for the SEC 3-statement model builder (``streamlit run app.py``)."""
import logging
from datetime import date

import pandas as pd
import streamlit as st

from sec_model import pipeline
from sec_model.checks import summarize
from sec_model.types import MarketInputs, UserFacingError

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
STATUS_COLORS = {"PASS": "#D9EAD3", "WARN": "#FFF2CC", "FAIL": "#F4CCCC", "SKIP": "#EEEEEE"}
STATEMENT_TABS = (("Income Statement", "IS"), ("Balance Sheet", "BS"), ("Cash Flow", "CF"))
UA_ERROR = "Enter a name and email for the SEC User-Agent."

log = logging.getLogger(__name__)


def statement_frame(stmt) -> pd.DataFrame:
    """One row per as-reported line: indented label, then a column per fiscal year
    (abstract/header rows carry no values)."""
    data = {"Line item": ["  " * r.indent + r.label for r in stmt.rows]}
    for fy in stmt.fy_labels:
        data[fy] = [None if r.is_abstract else r.values.get(fy) for r in stmt.rows]
    return pd.DataFrame(data)


def mapping_frame(mapping: dict, fy_labels: list[str]) -> pd.DataFrame:
    """One row per model concept: status, tag and label of the latest year, then values per FY."""
    latest = fy_labels[-1]
    rows = []
    for concept, by_fy in mapping.items():
        cv = by_fy.get(latest)
        row = {"Concept": concept, "Status": cv.status if cv else "UNMAPPED",
               "Tag": cv.tag if cv else None, "Label": cv.label if cv else None}
        for fy in fy_labels:
            fv = by_fy.get(fy)
            row[fy] = fv.value if fv else None
        rows.append(row)
    return pd.DataFrame(rows, columns=["Concept", "Status", "Tag", "Label", *fy_labels])


def checks_frame(checks, hide_skip: bool) -> pd.DataFrame:
    rows = [{"name": c.name, "fy": c.fy, "status": c.status, "lhs": c.lhs, "rhs": c.rhs,
             "tolerance": c.tolerance, "note": c.note}
            for c in checks if not (hide_skip and c.status == "SKIP")]
    return pd.DataFrame(rows, columns=["name", "fy", "status", "lhs", "rhs", "tolerance", "note"])


def _color_status(value) -> str:
    color = STATUS_COLORS.get(value)
    return f"background-color: {color}; color: #000000" if color else ""


def render_result(result) -> None:
    st.header(f"{result.company} ({result.ticker})")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("CIK", str(result.cik))
    c2.metric("Fiscal years", f"{result.fy_labels[0]}–{result.fy_labels[-1]}")
    c3.metric("Forecast", f"{result.fcst_labels[0]}–{result.fcst_labels[-1]}"
              if result.fcst_labels else "none")
    c4.metric("Currency", result.currency)
    c5.metric("Entity", result.entity_type.capitalize())

    st.subheader("Data checks")
    n = summarize(result.checks)
    st.markdown(f"**{len(result.checks)} checks:** " + " · ".join(
        f"{n.get(s, 0)} {s}" for s in ("PASS", "WARN", "FAIL", "SKIP")))
    hide_skip = st.checkbox("Hide SKIP rows", value=True, key="hide_skip")
    df = checks_frame(result.checks, hide_skip)
    st.dataframe(df.style.map(_color_status, subset=["status"]), hide_index=True,
                 width="stretch")
    for w in result.coverage:
        st.warning(w)

    st.download_button("Download Excel model", data=result.xlsx, mime=XLSX_MIME,
                       file_name=f"{result.ticker}_3Stmt_{date.today():%Y-%m-%d}.xlsx",
                       type="primary")

    tabs = st.tabs([*(label for label, _ in STATEMENT_TABS), "Mapping"])
    for tab, (_, stype) in zip(tabs, STATEMENT_TABS):
        with tab:
            st.dataframe(statement_frame(result.statements[stype]), hide_index=True,
                         width="stretch")
    with tabs[-1]:
        st.dataframe(mapping_frame(result.mapping, result.fy_labels), hide_index=True,
                     width="stretch")


def run_build(ticker: str, user_agent: str, n_years: int, n_fcst: int,
              market: MarketInputs, cik_override: int | None = None) -> None:
    """Build inside a live ``st.status`` and keep the result in session state."""
    with st.status("Building…", expanded=True) as status:
        try:
            result = pipeline.build(ticker, user_agent, n_years, n_fcst, market,
                                    progress=lambda m: st.write(m), cik_override=cik_override)
        except UserFacingError as e:
            status.update(label="Build failed", state="error")
            st.error(str(e))
            return
        except Exception:
            logging.exception("Unexpected error building %s", ticker)
            status.update(label="Build failed", state="error")
            st.error("Build failed; see server log.")
            return
        status.update(label="Done", state="complete", expanded=False)
        st.session_state["result"] = result


def main() -> None:
    st.set_page_config(page_title="SEC 3-Statement Model Builder", page_icon="📊", layout="wide")
    st.title("SEC 3-Statement Model Builder")
    st.caption("Pulls as-reported 10-K financials from SEC EDGAR, ties them out, and exports "
               "a linked 3-statement model with WACC and DCF — for US-GAAP 10-K filers.")

    with st.sidebar:
        st.header("Settings")
        ticker = st.text_input("Ticker", value="AAPL", key="ticker").strip().upper()
        user_agent = st.text_input("SEC User-Agent (name and email)", value="",
                                   placeholder="Jane Doe jane@example.com", key="ua").strip()
        n_years = st.slider("Years of history", 2, 15, 10, key="n_years")
        n_fcst = st.slider("Forecast years", 3, 7, 5, key="n_fcst")
        with st.expander("Market inputs"):
            price = st.number_input("Share / unit price", min_value=0.0, value=0.0,
                                    help="Leave 0 to enter later in the workbook", key="price")
            beta = st.number_input("Beta", value=1.0, key="beta")
            risk_free = st.number_input("Risk-free rate", value=0.0425, format="%.4f",
                                        key="risk_free")
            mrp = st.number_input("Market risk premium", value=0.055, format="%.4f", key="mrp")
        cik = st.number_input("CIK override (advanced)", min_value=0, value=0, step=1,
                              help="0 = none. Use when the ticker maps to a CIK without 10-K "
                                   "history (e.g. a new successor holding company).", key="cik")
        build_clicked = st.button("Build model", type="primary", width="stretch",
                                  key="build")
        st.caption("The SEC's fair-access policy requires requests to identify you: use your "
                   "real name and email. Requests are rate-limited and cached on disk.")

    if build_clicked:
        st.session_state.pop("result", None)
        if "@" not in user_agent:
            st.error(UA_ERROR)
        else:
            market = MarketInputs(price=price or None, beta=beta, risk_free=risk_free, mrp=mrp)
            run_build(ticker, user_agent, n_years, n_fcst, market, int(cik) or None)

    if "result" in st.session_state:
        render_result(st.session_state["result"])


if __name__ == "__main__":
    main()
