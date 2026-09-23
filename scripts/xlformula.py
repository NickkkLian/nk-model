#!/usr/bin/env python3
"""xlformula.py — read and evaluate the small part of Excel's formula language a three-statement model uses, the
way Excel evaluates it.

    python3 xlformula.py "SUM(B2:B4)*2" B2=1 B3=2 B4=3
    python3 xlformula.py --selftest

Numbers, text in double quotes, TRUE and FALSE; cell references (B5, $B$5, 'Cash flow'!B5) and ranges inside
SUM, MIN and MAX (B2:B7); + - * / ^ = <> < > <= >=; and SUM MIN MAX ABS IF AND OR. Excel's own rules, not a
calculator's: ^ runs left to right (2^3^2 is 64), a leading minus binds tighter than ^ (-2^2 is 4), SUM adds left to
right, a comparison between a number, a text and TRUE/FALSE is never equal (and orders number < text < logical), an
empty cell is 0, "" or FALSE to whatever it is compared with, and a step that has no finite answer is an error
(#DIV/0! for 1/0 and 0^-1, #NUM! for 0^0, for a negative number to a fractional power, and for anything too large).

Refused rather than guessed — the checker says it cannot verify these, and the models never use them: & (how Excel
writes a number as text was not tested here), ROUND (how Excel rounds 1.005, whose stored double is just below the
half, was not tested), TRUE, FALSE or text typed straight into SUM, MIN or MAX (Excel counts those, unlike the same
values in a range), and a range inside AND or OR (Excel skips its empty cells and text). Text in an IF's condition,
or typed into AND and OR, is #VALUE!. A formula outside this set is refused with its position.

Known to differ from Excel, and not refused (from Microsoft's documentation; none of it was tried in Excel here): a sum
or difference that comes out a hair from zero is kept as it is (1.333 + 1.225 - 1.333 - 1.225 gives -2.2e-16, where Excel 97
and later show 0); text in arithmetic is #VALUE! ("1"+"2" is 3 in Excel); text reached through a reference inside AND
or OR is #VALUE! (Excel ignores it). A TRUE or FALSE worked out inside SUM, MIN or MAX (SUM(1=1)) is skipped;
Microsoft's MAX page speaks only of values typed in and values in a range, and it was not tried. The model's own
formulas meet only the first; make_model.py writes the start cash and the check lines as 0 when they are within half a
penny.
"""
import math
import re
import sys

FUNCS = {"SUM": (1, None), "MIN": (1, None), "MAX": (1, None), "ABS": (1, 1), "IF": (2, 3), "AND": (1, None), "OR": (1, None)}
TOKEN = re.compile(r"""\s*(?:
    (?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Za-z0-9_(])
  | (?P<name>[A-Za-z_][A-Za-z0-9_.]*)
  | (?P<op><>|<=|>=|[-+*/^&=<>(),])
)""", re.X)


class FormulaError(ValueError):
    pass


def tokens(src):
    out, pos = [], 0
    src = src.rstrip()
    while pos < len(src):
        m = TOKEN.match(src, pos)
        if not m or m.end() == pos:
            raise FormulaError(f"cannot read the formula at {pos + 1}: {src[pos:pos + 12]!r}")
        kind = m.lastgroup
        out.append((kind, m.group(kind)))
        pos = m.end()
    return out


def split_sheet(text, here):
    """'Cash flow'!B5 -> ('Cash flow', 'B5'); B5 -> (here, 'B5')"""
    if "!" in text:
        sheet, cell = text.rsplit("!", 1)
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        return sheet, cell
    return here, text


def parse(src):
    """the formula (without its '=') as a tree"""
    toks, pos = tokens(src), [0]

    def peek():
        return toks[pos[0]] if pos[0] < len(toks) else ("end", None)

    def take(kind=None, value=None):
        t = peek()
        if (kind and t[0] != kind) or (value is not None and t[1] != value):
            raise FormulaError(f"expected {value or kind} but found {t[1] or 'the end'}")
        pos[0] += 1
        return t

    def is_op(v):
        return peek() == ("op", v)

    def compare():
        n = concat()
        while peek()[0] == "op" and peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = take()[1]; n = ("bin", op, n, concat())
        return n

    def concat():
        n = additive()
        if is_op("&"):
            raise FormulaError("the & operator is not one this checker can verify (how Excel writes a number as text was not tested)")
        return n

    def additive():
        n = term()
        while is_op("+") or is_op("-"):
            op = take()[1]; n = ("bin", op, n, term())
        return n

    def term():
        n = power()
        while is_op("*") or is_op("/"):
            op = take()[1]; n = ("bin", op, n, power())
        return n

    def power():                       # Excel: left to right, and below unary minus
        n = unary()
        while is_op("^"):
            take(); n = ("bin", "^", n, unary())
        return n

    def unary():
        if is_op("-"):
            take(); return ("neg", unary())
        if is_op("+"):
            take(); return unary()
        return atom()

    def atom():
        kind, text = peek()
        if kind == "num":
            take(); return ("num", float(text))
        if kind == "str":
            take(); return ("str", text[1:-1].replace('""', '"'))
        if kind == "ref":
            take(); return ("range", text) if ":" in text.split("!")[-1] else ("ref", text)
        if kind == "name":
            take()
            up = text.upper()
            if up in ("TRUE", "FALSE") and not is_op("("):
                return ("bool", up == "TRUE")
            if not is_op("("):
                raise FormulaError(f"'{text}' is not something this checker can read (a defined name?)")
            if up not in FUNCS:
                raise FormulaError(f"the function {text} is not one this checker can verify ({', '.join(sorted(FUNCS))})")
            take("op", "(")
            args = []
            if not is_op(")"):
                args.append(compare())
                while is_op(","):
                    take(); args.append(compare())
            take("op", ")")
            lo, hi = FUNCS[up]
            if len(args) < lo or (hi is not None and len(args) > hi):
                raise FormulaError(f"{up} takes {lo}{'' if hi == lo else ' or more' if hi is None else f' to {hi}'} arguments, not {len(args)}")
            if up in ("SUM", "MIN", "MAX") and any(a[0] in ("bool", "str") for a in args):
                raise FormulaError(f"TRUE, FALSE or text typed straight into {up} is not something this checker can verify (Excel counts it, unlike the same value in a range)")
            if up in ("AND", "OR") and any(a[0] == "range" for a in args):
                raise FormulaError(f"a range inside {up} is not something this checker can verify (Excel skips its empty cells and text)")
            return ("call", up, args)
        if is_op("("):
            take(); n = compare(); take("op", ")"); return n
        raise FormulaError(f"expected a value but found {text or 'the end'}")

    tree = compare()
    if peek()[0] != "end":
        raise FormulaError(f"unexpected {peek()[1]!r} after a complete formula")
    return tree


def refs(tree, here):
    """every (sheet, cell) the formula reads, ranges expanded"""
    from ooxml import split_ref, ref as mkref
    out = []

    def walk(n):
        if n[0] == "ref":
            s, c = split_sheet(n[1], here); out.append((s, c.replace("$", "")))
        elif n[0] == "range":
            s, span = split_sheet(n[1], here)
            a, b = span.replace("$", "").split(":")
            (r0, c0), (r1, c1) = split_ref(a), split_ref(b)
            for r in range(min(r0, r1), max(r0, r1) + 1):
                for c in range(min(c0, c1), max(c0, c1) + 1):
                    out.append((s, mkref(r, c)))
        elif n[0] in ("bin",):
            walk(n[2]); walk(n[3])
        elif n[0] == "neg":
            walk(n[1])
        elif n[0] == "call":
            for a in n[2]:
                walk(a)
    walk(tree)
    return out


def kind(v):
    """Excel's order of types in a comparison: numbers, then text, then TRUE/FALSE"""
    return 2 if isinstance(v, bool) else 1 if isinstance(v, str) else 0


def compare(op, a, b):
    """a comparison the way Excel makes it: an empty cell is 0, "" or FALSE to match the other side; values of
    different types are never equal and order as number < text < logical; text ignores case"""
    blank = {0: 0.0, 1: "", 2: False}
    if a is None and b is None:
        a = b = 0.0
    elif a is None:
        a = blank[kind(b)]
    elif b is None:
        b = blank[kind(a)]
    ka, kb = kind(a), kind(b)
    if ka != kb:
        x, y = ka, kb
    elif ka == 1:
        x, y = a.lower(), b.lower()
    else:
        x, y = float(a), float(b)
    return {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]


def finite(x):
    if isinstance(x, float) and not math.isfinite(x):
        raise FormulaError("#NUM! (a number too large for Excel)")
    return x


def evaluate(tree, here, get):
    """get(sheet, cell) -> the cell's value (None for an empty cell)"""
    def num(v):
        if v is None:
            return 0.0
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        raise FormulaError(f"#VALUE! (text {v!r} where a number is needed)")

    def cells(n):
        if n[0] == "range":
            from ooxml import split_ref, ref as mkref
            s, span = split_sheet(n[1], here)
            a, b = span.replace("$", "").split(":")
            (r0, c0), (r1, c1) = split_ref(a), split_ref(b)
            return [get(s, mkref(r, c)) for r in range(min(r0, r1), max(r0, r1) + 1) for c in range(min(c0, c1), max(c0, c1) + 1)]
        return [ev(n)]

    def ev(n):
        k = n[0]
        if k in ("num", "str", "bool"):
            return n[1]
        if k == "ref":
            s, c = split_sheet(n[1], here)
            return get(s, c.replace("$", ""))
        if k == "range":
            raise FormulaError("a range outside a function")
        if k == "neg":
            return -num(ev(n[1]))
        if k == "bin":
            op, a, b = n[1], ev(n[2]), ev(n[3])
            if op in ("=", "<>", "<", ">", "<=", ">="):
                return compare(op, a, b)
            x, y = num(a), num(b)
            if op == "+":
                return finite(x + y)
            if op == "-":
                return finite(x - y)
            if op == "*":
                return finite(x * y)
            if op == "/":
                if y == 0:
                    raise FormulaError("#DIV/0!")
                return finite(x / y)
            if x == 0 and y == 0:
                raise FormulaError("#NUM! (0^0)")
            if x == 0 and y < 0:
                raise FormulaError("#DIV/0! (0 to a negative power)")
            if x < 0 and y != int(y):
                raise FormulaError("#NUM! (a negative number to a fractional power)")
            try:
                return finite(math.pow(x, y))
            except (OverflowError, ValueError):
                raise FormulaError("#NUM! (a number too large for Excel)")
        if k == "call":
            f, args = n[1], n[2]
            if f == "IF":
                cond = ev(args[0])
                if isinstance(cond, str):
                    raise FormulaError("#VALUE! (text as an IF's condition)")
                return ev(args[1]) if num(cond) != 0 else (ev(args[2]) if len(args) > 2 else False)
            if f in ("SUM", "MIN", "MAX"):
                # numbers only: TRUE/FALSE and text in a cell are skipped, as Excel skips them in a reference
                nums = [float(v) for a in args for v in cells(a) if isinstance(v, (int, float)) and not isinstance(v, bool)]
                if f == "SUM":
                    total = 0.0
                    for x in nums:                  # left to right, as Excel adds; Python's sum() compensates since 3.12
                        total = finite(total + x)
                    return total
                return (min(nums) if f == "MIN" else max(nums)) if nums else 0.0
            if f == "ABS":
                return abs(num(ev(args[0])))
            if f in ("AND", "OR"):
                vals = [ev(a) for a in args]
                if any(isinstance(v, str) for v in vals):
                    raise FormulaError(f"#VALUE! (text inside {f})")
                vals = [num(v) != 0 for v in vals if v is not None]     # an empty cell is skipped, as Excel does
                if not vals:
                    raise FormulaError(f"#VALUE! ({f} with nothing to test)")
                return all(vals) if f == "AND" else any(vals)
        raise FormulaError(f"cannot evaluate {k}")

    return ev(tree)


def selftest():
    ok = []

    def say(cond, text):
        ok.append(bool(cond)); print(f"  {'PASS' if cond else 'FAIL'}  {text}")

    def run(src, cells=None, here="S"):
        cells = cells or {}
        return evaluate(parse(src), here, lambda s, c: cells.get((s, c)))

    say(run("2^3^2") == 64 and run("-2^2") == 4 and run("2*3+4") == 10 and run("2*(3+4)") == 14, "Excel's order: ^ left to right, a leading minus before ^")
    say(run("SUM(B2:B4)", {("S", "B2"): 1, ("S", "B3"): 2, ("S", "B4"): 3}) == 6, "SUM over a range")
    say(run("'Cash flow'!B5+Checks!C2", {("Cash flow", "B5"): 1.5, ("Checks", "C2"): 2}) == 3.5, "references to other sheets, quoted and not")
    say(run("$B$2*B$3", {("S", "B2"): 2, ("S", "B3"): 5}) == 10, "absolute and mixed references")
    say(run("MAX(0,B2)*0.2", {("S", "B2"): -100}) == 0 and run("MIN(B2,50)", {("S", "B2"): 80}) == 50, "MAX and MIN")
    say(run('IF(ABS(B2)<0.005,"balanced","NOT balanced")', {("S", "B2"): 0.001}) == "balanced", "IF with text results")
    say(run("AND(B2=0,B3=0)", {("S", "B2"): 0, ("S", "B3"): 0.0}) is True, "AND and comparisons")
    say(run("B9+1") == 1, "an empty cell counts as 0")
    refused = []
    for bad in ("VLOOKUP(A1,B1:C3,2)", "ROUND(B2,2)", "Price*2", "A1:B2", "1/0", "SUM(", "B2 B3"):
        try:
            run(bad); refused.append(False)
        except FormulaError:
            refused.append(True)
    say(all(refused), "refused: an unknown function, ROUND, a defined name, a bare range, division by zero, an unfinished call, two values in a row")

    def gives_error(src, cells=None):
        try:
            run(src, cells); return False
        except FormulaError:
            return True
    say(all(gives_error(x) for x in ("0^-1", "0^0", "10^400", "(-8)^(1/3)", '1&2', '"a"&1', "MAX(TRUE,0)", 'SUM("5",1)', "AND(B2:B3)", 'IF("abc",1,2)')),
        "no answer rather than a guess: 0^-1, 0^0, 10^400, (-8)^(1/3), &, TRUE or text typed into MAX/SUM, a range in AND, text as IF's condition")
    say(run("TRUE=1") is False and run('1<"a"') is True and run('"a"<TRUE') is True and run('B9=""') is True and run("B9=0") is True
        and run("B9=FALSE") is True and run('"ABC"="abc"') is True, "comparisons as Excel makes them: types never equal, number < text < logical, an empty cell matches \"\", 0 and FALSE")
    big = {("S", "B2"): 1e16, ("S", "B3"): 1.0, ("S", "B4"): -1e16}
    say(run("SUM(B2:B4)", big) == 0.0 and run("B2+B3+B4", big) == 0.0, "SUM adds left to right, as + does (1e16 + 1 - 1e16 is 0 both ways)")
    say(run("MAX(B2:B3)", {("S", "B2"): True, ("S", "B3"): 0.5}) == 0.5 and run("AND(B2,B3)", {("S", "B2"): 1, ("S", "B3"): None}) is True,
        "in a reference, TRUE is skipped by MAX and an empty cell by AND, as Excel does")
    say(sorted(refs(parse("SUM(B2:C3)+'Cash flow'!D4"), "S")) == [("Cash flow", "D4"), ("S", "B2"), ("S", "B3"), ("S", "C2"), ("S", "C3")],
        "the cells a formula reads, ranges expanded")
    print(f"xlformula selftest: {sum(ok)}/{len(ok)} passed")
    return 0 if all(ok) else 2


def main(argv):
    if "--selftest" in argv:
        return selftest()
    if not argv:
        print(__doc__.strip().split("\n\n")[1]); return 2
    cells = {}
    for a in argv[1:]:
        k, v = a.split("=", 1)
        cells[("S", k)] = float(v)
    try:
        print(evaluate(parse(argv[0].lstrip("=")), "S", lambda s, c: cells.get((s, c))))
        return 0
    except FormulaError as e:
        print(f"cannot evaluate: {e}"); return 1


if __name__ == "__main__":
    sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
    sys.exit(main(sys.argv[1:]))
