# General-Purpose Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the midstream-first build into a sector-neutral teaching/research tool: entity-aware labels (shares vs units), a revolver with a commitment, per-year minimum cash and its own rate, cosmetic polish, the parked mapping minors, a clean repo and a short how-to README.

**Architecture:** No new modules. Changes land in `model_spec.py` (lines, drivers, scalars), `concepts.py`/`mapping.py` (two concepts, three guards), `pipeline.py` (entity detection), the workbook writers (labels, blank first-column cells, validations), `app.py` (neutral copy) and the README. The Python evaluator stays the reference for every Excel formula.

**Tech Stack:** unchanged (Python 3.11+, openpyxl, streamlit, pytest).

**Spec:** `docs/superpowers/specs/2026-10-22-three-statement-model-design.md` plus the decisions below (this plan is the spec addendum; conversation decisions on 2026-10-25).

## Global Constraints

- Everything in the 2026-10-22 plan's Global Constraints still binds (deps, SEC client, millions, Excel styling, no Data Tables, no `#REF!`, `UserFacingError` only).
- Excel == Python: every new Model line/driver renders via `to_excel` and evaluates via `forecast.compute`; the regression's `bal_check_max < 1e-6` and the COM check must still pass.
- No sector words in product code or UI copy: "midstream", "pipeline", "East Daley", "asset" (as a line label) are gone from `sec_model/`, `app.py`, README. Entity words (unit/share, distribution/dividend) come only from the label table keyed by `entity_type`.
- Commit messages end with `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`.

## Review Focus

1. A partnership (EPD) and a corporation (OKE) built from the same code show "Diluted units / Distribution per unit" vs "Diluted shares / Dividend per share" everywhere (Model, Drivers, WACC, DCF, app). Test in Task 2.
2. A forecast year whose revolver need exceeds the commitment shows the shortfall line > 0 and a red revolver cell, and the balance sheet still ties (the draw is capped; cash falls below minimum). Test in Task 1.
3. A filer with a tagged drawn revolver (`LineOfCredit`) starts the forecast revolver at that balance with term debt reduced by the same amount; total debt history is unchanged. Test in Task 1.
4. First history column shows blanks, not whole-balance artefacts, on lag-based lines; the Excel cells are blank and `compute` returns None for them. Test in Task 3.
5. A ticker with only a related-party AR member row AND a custom-tagged trade-receivables row maps Receivables to the trade row. Test in Task 3.

---

### Task 1: Revolver upgrades

**Files:** Modify `sec_model/model_spec.py`, `sec_model/concepts.py` (one concept), `sec_model/forecast.py` (only if `_capped` signature changes), `sec_model/workbook/inputs_sheets.py` (new driver/scalar rows render automatically; verify), `sec_model/workbook/model_sheet.py` (conditional format on the shortfall line). Tests: `tests/test_forecast.py`, `tests/test_mapping.py`, `tests/test_workbook_model.py`.

**Interfaces (exact):**
- Concept `RevolverDrawn` (BS scope): tags `LineOfCredit`, `LongTermLineOfCredit`, `LinesOfCreditCurrent`, `RevolvingCreditFacility`; label `revolv|line of credit`; no sum. UNMAPPED → 0.
- Lines: `REVOLVER` hist = concept `RevolverDrawn` (via hidden `REV_RPT`), fcst unchanged (`REVOLVER₋₁ + REVOLVER_NET`). `DEBT` hist = `TotalDebt − REV_RPT` (term debt excluding the drawn revolver) so `DEBT + REVOLVER` = reported total debt; forecast `DEBT = DEBT₋₁ + DEBT_NET` unchanged. `DEBT_NET` hist = `(DEBT + REVOLVER) − (DEBT₋₁ + REVOLVER₋₁)` so history financing still ties; `CF_OTHER_FIN` unchanged.
- `INTEREST` fcst = `Driver("INT_RATE") * L("DEBT",1) + Driver("REV_RATE") * L("REVOLVER",1)`; `INT_RATE_HIST` denominator = `AVERAGE(DEBT + REVOLVER, prior)`.
- Drivers (per-year, in this order after `DEBT_ISSUE`): `MIN_CASH` (money; default `L("CASH")` last year — replaces the scalar), `REV_RATE` (pct; default `Line("INT_RATE_HIST")`, note "revolver rate; defaults to the term-debt rate").
- Scalars: remove `MIN_CASH`; add `REV_CAP` (money; default `0`; note "revolver commitment; 0 = unlimited"), `GROWTH_CAP` (pct; default `0.15`; note "cap on default growth rates"). `_capped(e)` uses `Scalar("GROWTH_CAP")` instead of the constant.
- `REVOLVER_NET` fcst = `IF(REV_CAP=0, need, MIN(need, REV_CAP − REVOLVER₋₁))` where `need = MAX(MIN_CASH − CASH_PRE, −REVOLVER₋₁)`; with `MIN_CASH` now `Driver("MIN_CASH")`.
- New CHECK-section line `LIQ_SHORTFALL` ("Liquidity shortfall vs minimum cash"), hist blank, fcst = `MAX(0, Driver("MIN_CASH") − L("CASH"))`; conditional format red when > 0.5; note "cash below minimum: revolver commitment exhausted".
- Workbook: `REV_CAP` and `GROWTH_CAP` get list-free numeric inputs (yellow); `LIQ_SHORTFALL` row placed after `BAL_CHECK`.

- [ ] **Step 1: Failing tests** — `test_revolver_capped_by_commitment_and_shortfall_reported` (TOY with huge DIST, `REV_CAP=50`: revolver balance == 50, LIQ_SHORTFALL > 0, BAL_CHECK 0 every year); `test_revolver_unlimited_when_cap_zero` (shortfall 0, cash == MIN_CASH); `test_opening_revolver_from_history` (RevolverDrawn=30, TotalDebt=130 → DEBT hist 100, REVOLVER hist 30, DEBT_NET hist unchanged vs before, interest year-1 = INT_RATE×100 + REV_RATE×30); `test_min_cash_per_year_driver` (override MIN_CASH dict {y1: 10, y2: 40} → cash follows); `test_growth_cap_scalar` (GROWTH_CAP=0.05 caps REV_GROWTH default); mapping test for `RevolverDrawn`; workbook test that LIQ_SHORTFALL has the conditional format and the drivers sheet has MIN_CASH and REV_RATE rows with formulas.
- [ ] **Step 2: Run, expect FAIL** → **Step 3: Implement** → **Step 4: Run full suite `-W error`, expect PASS** → **Step 5: Commit** `feat: revolver commitment, per-year minimum cash, revolver rate, opening balance from history`

---

### Task 2: Entity-aware labels and neutral defaults

**Files:** Create `sec_model/labels.py`; modify `sec_model/pipeline.py`, `sec_model/model_spec.py` (labels come from `labels.py`), `sec_model/workbook/inputs_sheets.py`, `model_sheet.py`, `wacc_sheet.py`, `dcf_sheet.py`, `build.py`, `app.py`. Tests: `tests/test_labels.py`, `tests/test_pipeline.py`, `tests/test_workbook_model.py`, `tests/test_app.py`.

**Interfaces (exact):**
- `labels.py`: `ENTITY_TYPES = ("corporation", "partnership")`; `LABELS: dict[str, dict[str, str]]` keyed by label id → entity → text. Ids and texts: `units` → "Diluted shares" / "Diluted units"; `units_out` → "Shares outstanding" / "Units outstanding"; `per_unit` → "per share" / "per unit"; `dist` → "Dividends paid" / "Distributions paid"; `dpu` → "Dividend per share" / "Distribution per unit"; `dpu_growth` → "Dividend per share growth" / "Distribution per unit growth"; `ni_common` → "Net income to common" / "Net income to partners"; `equity` → "Shareholders' equity" / "Partners' capital"; `holder` → "shareholders" / "unitholders". `label(id, entity_type) -> str`.
- `detect_entity_type(mapping, latest_fy) -> str`: "partnership" when the latest-year `EquityCommon` tag is `us-gaap_PartnersCapital` or any of `DilutedUnits`/`Distributions` tags contains `Partnership` or `LimitedPartner`; else "corporation".
- `build_spec(da_basis, entity_type="corporation")` — line and driver labels use `label(...)`; keys unchanged. `BuildResult.entity_type: str`.
- Sheet rename: `Asset Inputs` → `Revenue Streams` (constant `STREAMS_SHEET`, block headers "REVENUE BY STREAM" / "OPERATING COST BY STREAM", row 1 "Total company (placeholder)", last row "Other / unallocated", subtitle "Enter revenue and operating cost by stream (segment, product line, region, asset …) for each forecast year. Row 1 grows the consolidated figures off the Drivers sheet; overwrite it with real streams."). `ModelLayout.assets_sheet` → `streams_sheet`; `Block` keys stay `ASSET_REV`/`ASSET_COST` (internal).
- Defaults: `JV_TOGGLE` default `0.0` (note "1 = include equity-method earnings in EBITDA (common for partnerships and JV-heavy companies)"); `GROWTH_CAP` default stays 0.15 (Task 1).
- App copy: caption "Pulls as-reported 10-K financials from SEC EDGAR, ties them out, and exports a linked 3-statement model with WACC and DCF — for any US filer."; tab "Revenue Streams" nowhere (tabs unchanged); no sector words. Metric "Entity" shows the detected type.
- WACC/DCF: "Share / unit price" → `label("per_unit")`-driven ("Share price" / "Unit price"); "Basic units (MM)" → "Basic shares (MM)" / "Basic units (MM)"; over/under label unchanged.

- [ ] **Step 1: Failing tests** — `test_detect_entity_type_partnership_and_corporation`; `test_spec_labels_follow_entity_type` (EPD-shaped → "Diluted units", OKE-shaped → "Diluted shares"); workbook test that the Model, Drivers, WACC and DCF sheets contain "Distribution per unit" for a partnership spec and "Dividend per share" for a corporation; `test_no_sector_words` (grep `sec_model/`, `app.py` for `midstream|pipeline|east daley|Asset Inputs` → none); app test asserting the Entity metric.
- [ ] **Step 2–4: FAIL → implement → full suite PASS** → **Step 5: Commit** `feat: entity-aware labels, Revenue Streams sheet, neutral defaults`

---

### Task 3: Polish and parked minors

**Files:** `sec_model/workbook/model_sheet.py`, `sec_model/forecast.py`, `sec_model/model_spec.py` (a `lag_blank_first: bool` flag on LineDef or an automatic rule), `sec_model/mapping.py`, `sec_model/concepts.py`, `sec_model/workbook/dcf_sheet.py`. Tests: `tests/test_forecast.py`, `tests/test_workbook_model.py`, `tests/test_mapping.py`.

**Interfaces (exact):**
- First-column rule: a line whose `hist` expression references any `Line(key, lag≥1)` (and is not a concept read) evaluates to `None` in the FIRST history column and its Excel cell is left blank. Applies automatically (detect via `iter_nodes`); `Hist(...)` averages skip it (already skip None). Lines affected: `DEBT_NET`, `CF_DWC`, `REV_GROWTH_HIST`, `DPU_GROWTH_HIST`, `DA_PCT_HIST` (ppe basis), `INT_RATE_HIST`, `CF_OTHER_FIN`, `CF_OTHER_OP` (because of CF_DWC), `FX_OTHER`? — no: FX_OTHER has no lag; leave. Totals that sum a blanked line (`OCF`, `FCF_FIN`) must still show the reported value in the first column: `OCF` and `FCF_FIN` hist in the first column fall back to `OCF_RPT`/`FCF_RPT` (`IF(first column, RPT, sum)` is not expressible → instead make `CF_OTHER_OP`/`CF_OTHER_FIN` first-column value = the reported total minus the non-lag components, i.e. the derived lines absorb the blank; implement by evaluating the lag reference as 0 for the *derived absorber* lines only). Simplest correct rule: blank only the RATIO lines and `DEBT_NET`/`CF_DWC`, and let `CF_OTHER_OP`/`CF_OTHER_FIN` keep absorbing (they then equal "reported total − components" which is the honest first-year statement). Write the rule down in the module docstring.
- Mapping: label fallback runs BEFORE a lone-member sum (members ≥2 required for the member path to pre-empt); coverage regex → `^(accounts |trade |notes )?(receivable|payable)s?\b` excluding `income tax|interest`; gated `LongTermInvestments`/`Investments` additionally require the row label NOT to match `marketable|securities|available.for.sale`.
- DCF: note on the EV/EBIT multiple row "EBIT includes equity-method earnings (reported basis)".

- [ ] **Step 1: Failing tests** — first-column blanks (`compute` returns None for `DEBT_NET`, `REV_GROWTH_HIST` in FY1; workbook cells empty; `OCF` FY1 equals `OCF_RPT`); lone-member vs trade-row test (Review Focus 5); coverage regex test ("Income taxes payable" does not trigger); securities exclusion test.
- [ ] **Step 2–4: FAIL → implement → full suite PASS** → **Step 5: Commit** `fix: first-year artefacts blanked; mapping guards for related-party rows, securities and coverage noise`

---

### Task 4: Repo cleanup and README

**Files:** Delete `sec_10k_app.py`, `SEC_10K_Builder.ipynb`, `SEC_10K_Builder_v3 (working).ipynb`, `streamlit_SEC_jupyter.ipynb`; rewrite `README.md`; add `.streamlit/config.toml` (`[browser] gatherUsageStats = false`); leave `docs/`, `.devcontainer/`, untracked notebooks alone.

**README (exact sections, ≤ 180 lines):** Title + one-paragraph what-it-is → **Quick start** (install, run locally, type a ticker + your name/email, Build, Download) → **What you get** (sheet-by-sheet, one line each, with the colour legend) → **How to use the model** (edit Revenue Streams rows; Drivers; the three switches: JV toggle, working-capital hold, revolver commitment; where to enter price/beta; reading the Checks sheet; repointing a Mapping key) → **Checks** (table of the seven checks and their levels) → **Limitations** (40-F/20-F, non-USD shown unconverted, banks, predecessor history, XOM-style CIK override, subtotal checks) → **For developers** (tests, `scripts/regress.py`, the design/plan docs) → **Data source and fair use**.

- [ ] **Step 1: Delete files, write README and config** → **Step 2: `pytest tests -W error` PASS, `streamlit run app.py` starts (AppTest)** → **Step 3: Commit** `chore: remove the legacy app and notebooks; README how-to for the general-purpose release`

---

### Task 5: Regression, Excel check, acceptance

- [ ] **Step 1:** `python scripts/regress.py --ua "<name email>" OKE EPD ET KMI WMB MPLX TRGP KNTK ENB AAPL CVX JPM TRP XOM --excel` — exit 0 (XOM KNOWN). Fix root causes with tests if anything regressed.
- [ ] **Step 2:** COM check on OKE (corporation) and EPD (partnership): zero error cells; labels per entity; set `REV_CAP` to a small number on EPD and confirm the shortfall line lights up and BAL_CHECK stays 0; Drivers year-1 formulas == Python defaults.
- [ ] **Step 3:** Update `docs/superpowers/plans/2026-10-22-acceptance-run.md` with a "2026-10-25 general-purpose release" section (table + entity type column). Commit `test: acceptance run for the general-purpose release`.
