import pytest

from sec_model.formula import (
    Block, Const, Driver, EvalContext, Fn, Hist, Line, RenderContext, Scalar,
    evaluate, to_excel,
)

L = "ABCDEFGH"


def _a(c):
    return L[c - 1]


def make_ctx(col=4, first_col=1):
    return RenderContext(
        col=col,
        first_col=first_col,
        line_cell=lambda k, c: f"{_a(c)}{ {'REV': 5, 'COST': 6, 'X': 7}.get(k, 1) }",
        driver_cell=lambda k, c: f"Drivers!{_a(c)}{ {'MARGIN': 9}.get(k, 1) }",
        scalar_cell=lambda k: f"Inputs!$B${ {'MIN_CASH': 3, 'FLAG': 4}.get(k, 1) }",
        block_range=lambda k, c: f"SUM(Rev!{_a(c)}2:{_a(c)}4)",
        hist_range=lambda k, n: "$D$5:$F$5",
    )


def make_env(col=4, first_col=1, **vals):
    data = {"REV": 100.0, "COST": 60.0, "X": 0.0}
    data.update(vals)
    return EvalContext(
        col=col,
        first_col=first_col,
        line=lambda k, c: data.get(k) if c == col else data.get(k + "@" + str(c)),
        driver=lambda k, c: {"MARGIN": 0.5}.get(k),
        scalar=lambda k: {"MIN_CASH": 10.0, "FLAG": 1.0}.get(k),
        block=lambda k, c: 7.0,
        hist=lambda k, n: [1.0, None, 3.0],
    )


def test_arithmetic_renders_and_evaluates():
    e = (Line("REV") - Line("COST")) * Driver("MARGIN")
    assert to_excel(e, make_ctx()) == "=(D5-D6)*Drivers!D9"
    assert evaluate(e, make_env()) == pytest.approx((100 - 60) * 0.5)


def test_lag_resolves_prior_column_and_before_first_is_zero():
    e = Line("REV") - Line("REV", lag=1)
    assert to_excel(e, make_ctx(col=4)) == "=D5-C5"
    env = make_env(**{"REV@3": 80.0})
    assert evaluate(e, env) == pytest.approx(20.0)
    assert to_excel(e, make_ctx(col=3, first_col=3)) == "=C5-0"
    assert evaluate(e, make_env(col=3, first_col=3)) == pytest.approx(100.0)


def test_div_by_zero_is_zero_and_iferror_in_excel():
    e = Line("REV") / Line("X")
    assert to_excel(e, make_ctx()) == "=IFERROR(D5/D7,0)"
    assert evaluate(e, make_env()) == 0.0
    assert evaluate(Line("REV") / Line("COST"), make_env()) == pytest.approx(100 / 60)


def test_fn_max_if_and_block():
    e = Fn("MAX", Block("ASSETS"), Scalar("MIN_CASH"))
    assert to_excel(e, make_ctx()) == "=MAX(SUM(Rev!D2:D4),Inputs!$B$3)"
    assert evaluate(e, make_env()) == 10.0
    assert evaluate(Fn("MIN", Block("ASSETS"), Scalar("MIN_CASH")), make_env()) == 7.0
    cond = Scalar("FLAG").eq(1)
    f = Fn("IF", cond, Line("REV"), Line("COST"))
    assert to_excel(f, make_ctx()) == "=IF((Inputs!$B$4=1),D5,D6)"
    assert evaluate(f, make_env()) == 100.0
    g = Fn("IF", Scalar("FLAG").eq(0), Line("REV"), Line("COST"))
    assert evaluate(g, make_env()) == 60.0


def test_hist_average_range():
    e = Hist("REV", 3)
    assert to_excel(e, make_ctx()) == "=AVERAGE($D$5:$F$5)"
    assert evaluate(e, make_env()) == pytest.approx(2.0)  # None skipped
    assert evaluate(Hist("REV", 3), EvalContext(
        col=4, first_col=1, line=lambda k, c: 0.0, driver=lambda k, c: 0.0,
        scalar=lambda k: 0.0, block=lambda k, c: 0.0, hist=lambda k, n: [])) == 0.0


def test_neg_rendering_and_evaluation():
    assert to_excel(-Line("REV"), make_ctx()) == "=-D5"
    assert to_excel(-(Line("REV") + Line("COST")), make_ctx()) == "=-(D5+D6)"
    assert evaluate(-(Line("REV") + Line("COST")), make_env()) == -160.0


def test_const_rendering_and_int_autowrap():
    ctx = make_ctx()
    assert to_excel(Const(0.35), ctx) == "=0.35"
    assert to_excel(Const(2.0), ctx) == "=2"
    assert to_excel(Const(2), ctx) == "=2"
    e = 1 - Driver("MARGIN") * 2
    assert to_excel(e, ctx) == "=1-(Drivers!D9*2)"
    assert evaluate(e, make_env()) == pytest.approx(0.0)
    assert evaluate(2 / Line("COST"), make_env()) == pytest.approx(2 / 60)


def test_nested_binops_are_parenthesised():
    e = Line("REV") - (Line("COST") - Line("X"))
    assert to_excel(e, make_ctx()) == "=D5-(D6-D7)"
    assert evaluate(e, make_env()) == 40.0


def test_comparisons_evaluate_to_one_or_zero():
    env = make_env()
    r, c = Line("REV"), Line("COST")
    assert evaluate(r.gt(c), env) == 1.0
    assert evaluate(r.lt(c), env) == 0.0
    assert evaluate(r.ge(100), env) == 1.0
    assert evaluate(r.le(99), env) == 0.0
    assert evaluate(r.ne(c), env) == 1.0
    assert evaluate(r.eq(c), env) == 0.0
    assert to_excel(r.ne(c), make_ctx()) == "=(D5<>D6)"
    assert to_excel(r.ge(100), make_ctx()) == "=(D5>=100)"


def test_none_from_context_is_zero():
    env = make_env()
    assert evaluate(Driver("MISSING") + Scalar("MISSING") + Line("NOPE"), env) == 0.0
    assert evaluate(Line("REV") / Driver("MISSING"), env) == 0.0


def test_other_functions():
    env, ctx = make_env(), make_ctx()
    assert evaluate(Fn("ABS", Line("X") - Line("REV")), env) == 100.0
    assert evaluate(Fn("IFERROR", Line("REV"), Const(0)), env) == 100.0
    assert evaluate(Fn("AVERAGE", Line("REV"), Line("COST")), env) == 80.0
    assert to_excel(Fn("abs", Line("REV")), ctx) == "=ABS(D5)"


def test_unknown_function_rejected():
    with pytest.raises(ValueError):
        Fn("SUMPRODUCT", Const(1))


def test_fn_arity_checked():
    with pytest.raises(ValueError):
        Fn("IF", Const(1), Const(2))
