"""Evaluate a `ModelSpec` in Python: history columns, driver/scalar defaults, forecast.

This is the same model the workbook renders as Excel formulas; computing it here
lets tests prove the forecast balances. Columns are ``fy_labels + fcst_labels``;
column i (0-based) is formula column ``FIRST_COL + i``.

Order of work:
1. History columns, each in topological order of its lag-0 `Line` references
   (`eval_order(spec, "hist")`). A concept-name `hist` reads
   ``history[concept][fy]``; an unmapped concept-year stays None (a blank cell).
   A `blank_first` line is None in the first history column (`model_spec` docstring).
   Only scalars with a constant default (or an override) are visible here.
2. Driver and scalar defaults, evaluated in the last history column.
3. Driver values per forecast year: the default (or an override) in the first
   year, later years chain the previous year's value, as Excel's ``=prev`` does.
4. Placeholder revenue-stream block, unless one is given.
5. Forecast columns, each in topological order (`eval_order(spec, "fcst")`).

A line with no formula for a column range is None there (blank in Excel).
"""
from __future__ import annotations

from typing import Iterator

from .formula import (
    Block, Driver, EvalContext, Expr, Fn, Hist, Line, Scalar, evaluate, iter_nodes,
)
from .model_spec import BLOCKS, ModelSpec

FIRST_COL = 1
History = dict[str, dict[str, "float | None"]]


def _formula(ld, phase: str):
    return ld.hist if phase == "hist" else ld.fcst


def _exprs(spec: ModelSpec) -> Iterator[tuple[str, Expr]]:
    for ld in spec.lines:
        for phase in ("hist", "fcst"):
            f = _formula(ld, phase)
            if isinstance(f, Expr):
                yield f"{ld.key}.{phase}", f
    for d in spec.drivers:
        yield f"driver {d.key}", d.default
    for s in spec.scalars:
        if isinstance(s.default, Expr):
            yield f"scalar {s.key}", s.default


def validate_spec(spec: ModelSpec) -> None:
    """Raise ValueError on duplicate keys, unknown references or bad IFERROR arity."""
    lines = spec.line_keys()
    dupes = {k for k in lines if lines.count(k) > 1}
    if dupes:
        raise ValueError(f"duplicate line keys: {sorted(dupes)}")
    known = {Line: set(lines), Hist: set(lines), Driver: {d.key for d in spec.drivers},
             Scalar: {s.key for s in spec.scalars}, Block: set(BLOCKS)}
    for where, expr in _exprs(spec):
        for node in iter_nodes(expr):
            kind = type(node)
            if kind in known and node.key not in known[kind]:
                raise ValueError(f"{where}: unknown {kind.__name__} {node.key!r}")
            if isinstance(node, Fn) and node.name == "IFERROR" and len(node.args) != 2:
                raise ValueError(f"{where}: IFERROR needs exactly 2 arguments")


def eval_order(spec: ModelSpec, phase: str) -> list[str]:
    """Topological order of lines within one column for phase "hist" or "fcst".

    Edges are lag-0 `Line` references; lag >= 1 points at an earlier column. Raises
    ValueError if the same-column graph has a cycle (it would be circular in Excel).
    """
    deps = {}
    for ld in spec.lines:
        f = _formula(ld, phase)
        deps[ld.key] = ({n.key for n in iter_nodes(f) if isinstance(n, Line) and n.lag == 0}
                        if isinstance(f, Expr) else set())
    order, done, pending = [], set(), list(deps)
    while pending:
        ready = [k for k in pending if deps[k] <= done]
        if not ready:
            raise ValueError(f"circular same-column references in {phase} formulas: {pending}")
        for k in ready:
            order.append(k)
            done.add(k)
        pending = [k for k in pending if k not in done]
    return order


def _no_driver(key, col):
    raise ValueError(f"driver {key!r} is not available in history columns")


def _no_block(key, col):
    raise ValueError(f"block {key!r} is not available in history columns")


class _Model:
    """Holds computed values and builds EvalContexts over them."""

    def __init__(self, spec: ModelSpec, fy_labels: list[str], fcst_labels: list[str]):
        if not fy_labels:
            raise ValueError("at least one history year is required")
        validate_spec(spec)
        self.spec = spec
        self.defs = {ld.key: ld for ld in spec.lines}
        self.fy = list(fy_labels)
        self.cols = self.fy + list(fcst_labels)
        self.vals: dict[str, dict[str, float | None]] = {k: {} for k in spec.line_keys()}

    def label(self, col: int) -> str:
        return self.cols[col - FIRST_COL]

    def env(self, col: int, driver=_no_driver, scalar=None, block=_no_block) -> EvalContext:
        return EvalContext(
            col=col, first_col=FIRST_COL,
            line=lambda k, c: self.vals[k][self.label(c)],
            driver=driver, scalar=scalar, block=block,
            hist=lambda k, n: [self.vals[k][fy] for fy in self.fy[-n:]],
        )

    def run_history(self, history: History, scalar) -> None:
        order = eval_order(self.spec, "hist")
        for i, fy in enumerate(self.fy):
            env = self.env(FIRST_COL + i, scalar=scalar)
            for key in order:
                f = self.defs[key].hist
                if f is None or (i == 0 and self.defs[key].blank_first):
                    v = None
                elif isinstance(f, str):
                    raw = (history.get(f) or {}).get(fy)
                    v = None if raw is None else float(raw)
                else:
                    v = evaluate(f, env)
                self.vals[key][fy] = v

    def last_hist_env(self, scalar) -> EvalContext:
        return self.env(FIRST_COL + len(self.fy) - 1, scalar=scalar)


def _history_scalar(spec: ModelSpec, overrides: dict):
    """Scalar lookup for history columns: overrides or constant defaults only."""
    consts = {s.key: float(s.default) for s in spec.scalars if not isinstance(s.default, Expr)}
    consts.update({k: float(v) for k, v in overrides.items()})

    def scalar(key):
        if key not in consts:
            raise ValueError(f"scalar {key!r} depends on history; not usable in history")
        return consts[key]
    return scalar


def _check_keys(overrides: dict, valid, what: str) -> None:
    unknown = set(overrides) - set(valid)
    if unknown:
        raise KeyError(f"unknown {what} override(s): {sorted(unknown)}")


def _prepare(spec, history, fy_labels, fcst_labels, scalar_overrides):
    scalar_overrides = dict(scalar_overrides or {})
    _check_keys(scalar_overrides, [s.key for s in spec.scalars], "scalar")
    m = _Model(spec, fy_labels, fcst_labels)
    hist_scalar = _history_scalar(spec, scalar_overrides)
    m.run_history(history, hist_scalar)
    env = m.last_hist_env(hist_scalar)
    scalars = {s.key: float(evaluate(s.default, env)) if isinstance(s.default, Expr)
               else float(s.default) for s in spec.scalars}
    scalars.update({k: float(v) for k, v in scalar_overrides.items()})
    env = m.last_hist_env(lambda k: scalars[k])
    drivers = {d.key: evaluate(d.default, env) for d in spec.drivers}
    return m, drivers, scalars


def driver_defaults(spec: ModelSpec, history: History, fy_labels: list[str]) -> dict[str, float]:
    """Each driver's default value (its first forecast column), from history."""
    return _prepare(spec, history, fy_labels, [], None)[1]


def scalar_defaults(spec: ModelSpec, history: History, fy_labels: list[str]) -> dict[str, float]:
    """Each scalar's default value; Expr defaults evaluate in the last history column."""
    return _prepare(spec, history, fy_labels, [], None)[2]


def _driver_series(defaults: dict, overrides: dict, fcst_labels: list[str]) -> dict:
    """Per-year driver values: an override is a float (all years) or {year: value};
    years without a value chain the previous year's, as the Drivers sheet does."""
    out = {}
    for key, default in defaults.items():
        ov = overrides.get(key)
        v = float(ov) if isinstance(ov, (int, float)) else default
        per_year = ov if isinstance(ov, dict) else {}
        _check_keys(per_year, fcst_labels, f"{key} year")
        series = {}
        for fy in fcst_labels:
            v = float(per_year.get(fy, v))
            series[fy] = v
        out[key] = series
    return out


def _default_asset_block(rev_last: float, drivers: dict, fcst_labels: list[str]) -> dict:
    rev, block = rev_last, {"ASSET_REV": {}, "ASSET_COST": {}}
    for fy in fcst_labels:
        rev = rev * (1 + drivers["REV_GROWTH"][fy])
        block["ASSET_REV"][fy] = rev
        block["ASSET_COST"][fy] = drivers["COST_PCT"][fy] * rev
    return block


def compute(spec: ModelSpec, history: History, fy_labels: list[str], fcst_labels: list[str],
            driver_overrides: dict | None = None, scalar_overrides: dict | None = None,
            asset_block: dict | None = None) -> dict[str, dict[str, float | None]]:
    """Every line x every column (history then forecast). Blank cells are None.

    driver_overrides: {driver: float | {fcst_year: float}}; scalar_overrides:
    {scalar: float}; asset_block: {"ASSET_REV"|"ASSET_COST": {fcst_year: float}},
    default one placeholder stream growing last revenue at REV_GROWTH with cost at COST_PCT.
    """
    driver_overrides = dict(driver_overrides or {})
    _check_keys(driver_overrides, [d.key for d in spec.drivers], "driver")
    m, defaults, scalars = _prepare(spec, history, fy_labels, fcst_labels, scalar_overrides)
    drivers = _driver_series(defaults, driver_overrides, list(fcst_labels))
    if asset_block is None:
        rev_last = m.vals["REV"][fy_labels[-1]] if "REV" in m.vals else None
        asset_block = _default_asset_block(rev_last or 0.0, drivers, list(fcst_labels))

    order = eval_order(spec, "fcst")
    for i, fy in enumerate(fcst_labels):
        env = m.env(
            FIRST_COL + len(fy_labels) + i,
            driver=lambda k, c: drivers[k][m.label(c)],
            scalar=lambda k: scalars[k],
            block=lambda k, c: (asset_block.get(k) or {}).get(m.label(c)),
        )
        for key in order:
            f = m.defs[key].fcst
            m.vals[key][fy] = None if f is None else evaluate(f, env)
    return m.vals
