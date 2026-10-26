"""Streamlit app wiring, driven headlessly with AppTest (fake build, plus one real OKE run)."""
from datetime import date
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from sec_model.pipeline import BuildResult
from sec_model.types import (CheckResult, ConceptValue, MarketInputs, MergedRow,
                             MergedStatement, UserFacingError)

ROOT = Path(__file__).resolve().parent.parent
APP = str(ROOT / "app.py")
UA = "SEC-Model-Framework dev admin@secmodelframework.dev"
FYS = ["FY2024", "FY2025"]


def _stmt(stype):
    rows = [MergedRow("hdr", "us-gaap_Hdr", "Header", 0, True, "money", {}, True),
            MergedRow("us-gaap_X", "us-gaap_X", "Child line", 2, False, "money",
                      {"FY2024": 1.5, "FY2025": 2.5}, True)]
    return MergedStatement(stype, "USD", list(FYS), {}, rows, {})


def fake_result(ticker="OKE"):
    checks = [CheckResult("Balance", "FY2025", "PASS", 1.0, 1.0, 0.5, ""),
              CheckResult("Cash roll", "FY2025", "WARN", 1.0, 2.0, 0.5, "off"),
              CheckResult("Subtotals", "FY2025", "SKIP", None, None, 0.5, "n/a")]
    mapping = {"Revenue": {fy: ConceptValue("Revenue", fy, 10.0, "k", "us-gaap_Revenues",
                                            "Total revenues", "OK") for fy in FYS}}
    return BuildResult(
        company="ONEOK INC", ticker=ticker, cik=1039684, currency="USD", fy_labels=list(FYS),
        fcst_labels=["FY2026E", "FY2027E", "FY2028E"],
        statements={s: _stmt(s) for s in ("IS", "BS", "CF")}, mapping=mapping, checks=checks,
        coverage=["Capex is unmapped"], forecast={}, xlsx=b"PK\x03\x04fake", log=[],
        entity_type="partnership")


@pytest.fixture
def calls(monkeypatch):
    seen = []

    def fake_build(ticker, user_agent, n_years=10, n_fcst=5, market=MarketInputs(),
                   progress=print, *a, **k):
        seen.append((ticker, user_agent, n_years, n_fcst, market, k.get("cik_override")))
        progress("step one")
        return fake_result(ticker)

    monkeypatch.setattr("sec_model.pipeline.build", fake_build)
    return seen


def _fresh():
    return AppTest.from_file(APP, default_timeout=60).run()


def _download_buttons(at):
    return at.get("download_button")


def test_initial_render_has_all_inputs():
    at = _fresh()
    assert not at.exception
    assert at.title[0].value == "SEC 3-Statement Model Builder"
    assert at.text_input(key="ticker").value == "AAPL"
    assert at.caption[0].value.endswith("— for US-GAAP 10-K filers.")
    assert at.text_input(key="ua").value == ""
    assert at.text_input(key="ua").placeholder == "Jane Doe jane@example.com"
    assert at.slider(key="n_years").value == 10 and at.slider(key="n_fcst").value == 5
    assert at.number_input(key="price").value == 0.0
    assert at.number_input(key="risk_free").value == pytest.approx(0.0425)
    assert at.button(key="build").label == "Build model"
    assert not _download_buttons(at)


def test_blank_user_agent_is_rejected_without_building(calls):
    at = _fresh()
    at.text_input(key="ua").set_value("not an email")
    at.button(key="build").click().run()
    assert [e.value for e in at.error] == ["Enter a name and email for the SEC User-Agent."]
    assert not calls and not at.exception


def test_successful_build_shows_results_and_download(calls):
    at = _fresh()
    at.text_input(key="ticker").set_value(" oke ")
    at.text_input(key="ua").set_value(UA)
    at.number_input(key="price").set_value(55.0)
    at.button(key="build").click().run()
    assert not at.exception and not at.error
    assert calls == [("OKE", UA, 10, 5, MarketInputs(price=55.0), None)]
    assert "ONEOK INC (OKE)" in [h.value for h in at.header]
    metrics = {m.label: m.value for m in at.metric}
    assert metrics == {"CIK": "1039684", "Fiscal years": "FY2024–FY2025",
                       "Forecast": "FY2026E–FY2028E", "Currency": "USD",
                       "Entity": "Partnership"}
    assert [w.value for w in at.warning] == ["Capex is unmapped"]
    assert at.status[0].label == "Done" and at.status[0].state == "complete"
    (dl,) = _download_buttons(at)
    assert dl.proto.label == "Download Excel model"
    assert [t.label for t in at.tabs] == ["Income Statement", "Balance Sheet", "Cash Flow",
                                          "Mapping"]
    # checks table, SKIP hidden by default, then shown
    checks_df = at.dataframe[0].value
    assert list(checks_df.columns) == ["name", "fy", "status", "lhs", "rhs", "tolerance", "note"]
    assert set(checks_df["status"]) == {"PASS", "WARN"}
    at.checkbox(key="hide_skip").uncheck().run()
    assert set(at.dataframe[0].value["status"]) == {"PASS", "WARN", "SKIP"}
    assert any("1 PASS · 1 WARN · 0 FAIL · 1 SKIP" in m.value for m in at.markdown)


def test_zero_price_means_blank(calls):
    at = _fresh()
    at.text_input(key="ua").set_value(UA)
    at.button(key="build").click().run()
    assert calls[0][4].price is None


def test_cik_override_zero_means_none_and_value_is_passed(calls):
    at = _fresh()
    assert at.number_input(key="cik").value == 0
    at.text_input(key="ua").set_value(UA)
    at.button(key="build").click().run()
    at.number_input(key="cik").set_value(34088)
    at.button(key="build").click().run()
    assert calls[0][5] is None and calls[1][5] == 34088


def test_user_facing_error_is_shown_verbatim(monkeypatch):
    def boom(*a, **k):
        raise UserFacingError("Ticker ZZZZ not found.")
    monkeypatch.setattr("sec_model.pipeline.build", boom)
    at = _fresh()
    at.text_input(key="ua").set_value(UA)
    at.button(key="build").click().run()
    assert [e.value for e in at.error] == ["Ticker ZZZZ not found."]
    assert at.status[0].label == "Build failed" and at.status[0].state == "error"
    assert not _download_buttons(at) and not at.exception


def test_unexpected_error_hides_traceback_and_is_logged(monkeypatch, caplog):
    def boom(*a, **k):
        raise RuntimeError("secret internals")
    monkeypatch.setattr("sec_model.pipeline.build", boom)
    at = _fresh()
    at.text_input(key="ua").set_value(UA)
    at.button(key="build").click().run()
    assert [e.value for e in at.error] == ["Build failed; see server log."]
    assert not at.exception
    assert "secret internals" in caplog.text


def test_shim_runs_the_app():
    at = AppTest.from_file(str(ROOT / "streamlit_app.py"), default_timeout=60).run()
    assert not at.exception
    assert at.title[0].value == "SEC 3-Statement Model Builder"


@pytest.mark.network
@pytest.mark.skipif(not (ROOT / ".cache").is_dir(), reason="SEC cache not present")
def test_real_oke_build_through_the_app():
    at = AppTest.from_file(APP, default_timeout=180).run()
    at.text_input(key="ticker").set_value("OKE")
    at.text_input(key="ua").set_value(UA)
    at.button(key="build").click().run()
    assert not at.exception and not at.error
    assert at.status[0].state == "complete"
    (dl,) = _download_buttons(at)
    assert dl.proto.label == "Download Excel model"
    metrics = {m.label: m.value for m in at.metric}
    assert (metrics["Currency"], metrics["Entity"]) == ("USD", "Corporation")
    assert len(at.tabs) == 4
