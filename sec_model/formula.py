"""Formula AST: one definition, two renderers (Excel string, Python value).

The forecast model is written once as expressions over named lines/drivers.
`to_excel` renders an expression to an Excel formula for one worksheet column;
`evaluate` computes the same expression in Python. Columns are 1-based worksheet
column indexes; `Line(key, lag)` resolves at ``col - lag``.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Iterator

FUNCTIONS = ("MAX", "MIN", "IF", "AVERAGE", "ABS", "IFERROR")
_CMP_OPS = ("=", "<>", "<", "<=", ">", ">=")


def _wrap(x) -> "Expr":
    if isinstance(x, Expr):
        return x
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return Const(float(x))
    raise TypeError(f"cannot use {type(x).__name__} in a formula expression")


class Expr:
    """Base node. Arithmetic operators build BinOp/Neg; comparisons via eq/ne/lt/..."""

    def __add__(self, o): return BinOp("+", self, _wrap(o))
    def __radd__(self, o): return BinOp("+", _wrap(o), self)
    def __sub__(self, o): return BinOp("-", self, _wrap(o))
    def __rsub__(self, o): return BinOp("-", _wrap(o), self)
    def __mul__(self, o): return BinOp("*", self, _wrap(o))
    def __rmul__(self, o): return BinOp("*", _wrap(o), self)
    def __truediv__(self, o): return BinOp("/", self, _wrap(o))
    def __rtruediv__(self, o): return BinOp("/", _wrap(o), self)
    def __neg__(self): return Neg(self)

    def eq(self, o): return Cmp("=", self, _wrap(o))
    def ne(self, o): return Cmp("<>", self, _wrap(o))
    def lt(self, o): return Cmp("<", self, _wrap(o))
    def le(self, o): return Cmp("<=", self, _wrap(o))
    def gt(self, o): return Cmp(">", self, _wrap(o))
    def ge(self, o): return Cmp(">=", self, _wrap(o))


@dataclass(frozen=True)
class Const(Expr):
    value: float


@dataclass(frozen=True)
class Line(Expr):
    key: str
    lag: int = 0


@dataclass(frozen=True)
class Driver(Expr):
    key: str


@dataclass(frozen=True)
class Scalar(Expr):
    key: str


@dataclass(frozen=True)
class Block(Expr):
    key: str


@dataclass(frozen=True)
class Hist(Expr):
    key: str
    n: int


@dataclass(frozen=True)
class BinOp(Expr):
    op: str
    left: Expr
    right: Expr


@dataclass(frozen=True)
class Neg(Expr):
    operand: Expr


@dataclass(frozen=True)
class Cmp(Expr):
    op: str
    left: Expr
    right: Expr

    def __post_init__(self):
        if self.op not in _CMP_OPS:
            raise ValueError(f"unsupported comparison {self.op!r}")


@dataclass(frozen=True, init=False)
class Fn(Expr):
    name: str
    args: tuple

    def __init__(self, name: str, *args):
        name = name.upper()
        if name not in FUNCTIONS:
            raise ValueError(f"unsupported function {name!r}")
        if (name == "IF" and len(args) != 3) or (name == "ABS" and len(args) != 1)                 or not args:
            raise ValueError(f"wrong number of arguments for {name}")
        object.__setattr__(self, "name", name)
        object.__setattr__(self, "args", tuple(_wrap(a) for a in args))


def iter_nodes(expr: Expr) -> Iterator[Expr]:
    """Yield every node of an expression tree (pre-order)."""
    yield expr
    if isinstance(expr, (BinOp, Cmp)):
        yield from iter_nodes(expr.left)
        yield from iter_nodes(expr.right)
    elif isinstance(expr, Neg):
        yield from iter_nodes(expr.operand)
    elif isinstance(expr, Fn):
        for a in expr.args:
            yield from iter_nodes(a)


@dataclass
class RenderContext:
    col: int                                    # 1-based column of the cell being rendered
    line_cell: Callable[[str, int], str]        # (key, col) -> "D12"
    driver_cell: Callable[[str, int], str]      # (key, col) -> "Drivers!D7"
    scalar_cell: Callable[[str], str]           # key -> "Inputs!$B$3"
    block_range: Callable[[str, int], str]      # (key, col) -> "SUM(Rev!D2:D9)"
    hist_range: Callable[[str, int], str]       # (key, n) -> "$D$5:$F$5"
    first_col: int = 1                          # first column that holds data


@dataclass
class EvalContext:
    col: int
    line: Callable[[str, int], float]           # (key, col) -> value
    driver: Callable[[str, int], float]         # (key, col) -> value
    scalar: Callable[[str], float]              # key -> value
    block: Callable[[str, int], float]          # (key, col) -> value
    hist: Callable[[str, int], list]            # (key, n) -> last n history values
    first_col: int = 1


def _const_text(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else repr(float(v))


def _operand(e: Expr, ctx: RenderContext) -> str:
    s = _render(e, ctx)
    return f"({s})" if isinstance(e, BinOp) else s


def _render(e: Expr, ctx: RenderContext) -> str:
    if isinstance(e, Const):
        return _const_text(e.value)
    if isinstance(e, Line):
        c = ctx.col - e.lag
        return "0" if c < ctx.first_col else ctx.line_cell(e.key, c)
    if isinstance(e, Driver):
        return ctx.driver_cell(e.key, ctx.col)
    if isinstance(e, Scalar):
        return ctx.scalar_cell(e.key)
    if isinstance(e, Block):
        return ctx.block_range(e.key, ctx.col)
    if isinstance(e, Hist):
        return f"AVERAGE({ctx.hist_range(e.key, e.n)})"
    if isinstance(e, Neg):
        return f"-{_operand(e.operand, ctx)}"
    if isinstance(e, Cmp):
        return f"({_render(e.left, ctx)}{e.op}{_render(e.right, ctx)})"
    if isinstance(e, Fn):
        return f"{e.name}({','.join(_render(a, ctx) for a in e.args)})"
    if isinstance(e, BinOp):
        a, b = _operand(e.left, ctx), _operand(e.right, ctx)
        return f"IFERROR({a}/{b},0)" if e.op == "/" else f"{a}{e.op}{b}"
    raise TypeError(f"unknown expression node {type(e).__name__}")


def to_excel(expr: Expr, ctx: RenderContext) -> str:
    """Render `expr` as an Excel formula string (leading '=') for column ctx.col."""
    return "=" + _render(_wrap(expr), ctx)


def _num(v) -> float:
    return 0.0 if v is None else float(v)


def _eval(e: Expr, env: EvalContext) -> float:
    if isinstance(e, Const):
        return float(e.value)
    if isinstance(e, Line):
        c = env.col - e.lag
        return 0.0 if c < env.first_col else _num(env.line(e.key, c))
    if isinstance(e, Driver):
        return _num(env.driver(e.key, env.col))
    if isinstance(e, Scalar):
        return _num(env.scalar(e.key))
    if isinstance(e, Block):
        return _num(env.block(e.key, env.col))
    if isinstance(e, Hist):
        vals = [float(v) for v in env.hist(e.key, e.n) if v is not None]
        return sum(vals) / len(vals) if vals else 0.0
    if isinstance(e, Neg):
        return -_eval(e.operand, env)
    if isinstance(e, Cmp):
        a, b = _eval(e.left, env), _eval(e.right, env)
        ok = {"=": a == b, "<>": a != b, "<": a < b, "<=": a <= b,
              ">": a > b, ">=": a >= b}[e.op]
        return 1.0 if ok else 0.0
    if isinstance(e, Fn):
        if e.name == "IF":
            cond, a, b = e.args
            return _eval(a if _eval(cond, env) != 0 else b, env)
        if e.name == "IFERROR":
            return _eval(e.args[0], env)
        vals = [_eval(a, env) for a in e.args]
        if e.name == "MAX":
            return max(vals)
        if e.name == "MIN":
            return min(vals)
        if e.name == "AVERAGE":
            return sum(vals) / len(vals)
        return abs(vals[0])  # ABS
    if isinstance(e, BinOp):
        a, b = _eval(e.left, env), _eval(e.right, env)
        if e.op == "+":
            return a + b
        if e.op == "-":
            return a - b
        if e.op == "*":
            return a * b
        return 0.0 if b == 0 else a / b
    raise TypeError(f"unknown expression node {type(e).__name__}")


def evaluate(expr: Expr, env: EvalContext) -> float:
    """Evaluate `expr` in Python at column env.col; None -> 0.0, x/0 -> 0.0."""
    return _eval(_wrap(expr), env)
