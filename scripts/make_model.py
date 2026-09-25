#!/usr/bin/env python3
"""make_model.py — turn a business, written down as a few assumptions, into a three-statement model in one .xlsx.

    python3 make_model.py business.json -o model.xlsx
    python3 make_model.py --init business.json          # a filled-in example to edit
    python3 make_model.py --selftest

The workbook has six sheets: Assumptions (every input, and how the model works), Income statement, Balance sheet,
Cash flow, Checks, and Charts. Every number on the three statements is a formula that points back at the assumptions
or at another statement, stored together with the value it gives, so the file reads correctly in Excel (which
recalculates) and in viewers that do not (which show the stored value).

Before it writes anything it checks the business (rules B01-B03 below), builds the workbook, fills in every formula's
value, and then works the whole model out a second time in plain Python, without the formulas. If the two disagree
anywhere, or the balance sheet does not balance to the penny in any year, it writes nothing and says where.
  B01  the file: valid JSON, every field there with the right type, no unknown or repeated keys; text with no control
       characters and no emoji or icon characters (model_check.py's K11 would refuse the workbook otherwise)
  B02  the numbers: in ranges that make sense for a small business (a price above 0, days between 0 and 365, a tax
       rate below 100%, no amount above a billion...)
  B03  honesty: made-up numbers are marked illustrative; real ones say where they come from
Exit 0 written · 1 refused, nothing written · 2 usage or selftest failed.
"""
import io, json, math, os, re, sys

sys.dont_write_bytecode = True
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import ooxml as O
import xlformula as X

# ---- the business file -------------------------------------------------------------------------------------------
SPEC = {
    "name": "str", "description": "str", "currency": "str", "first_year": "int", "years": "int", "illustrative": "bool",
    "sources": "str?",
    "sales": {"units": "num", "unit_growth": "num", "price": "num", "price_growth": "num"},
    "costs": {"unit_cost": "num", "unit_cost_growth": "num", "fixed_costs": "num", "fixed_cost_growth": "num"},
    "equipment": {"initial": "num", "yearly": "num", "depreciation_rate": "num"},
    "working_capital": {"receivable_days": "num", "inventory_days": "num", "payable_days": "num"},
    "funding": {"equity": "num", "loan": "num", "interest_rate": "num", "loan_years": "int"},
    "tax_rate": "num",
}
BILLION = 1e9     # a small business: no amount, count or price above a billion (this does not by itself keep every sum exact to the penny: SKILL.md, Boundaries)
RANGES = [   # (path, low, high, words) — inclusive
    ("years", 1, 10, "the number of years"), ("first_year", 1900, 2200, "the first year"),
    ("sales.units", 0, BILLION, "units sold"), ("sales.unit_growth", -0.9, 5, "the growth in units (a fraction: 0.12 is 12%)"),
    ("sales.price", 1e-9, BILLION, "the price"), ("sales.price_growth", -0.9, 5, "the growth in price (a fraction)"),
    ("costs.unit_cost", 0, BILLION, "the cost of a unit"), ("costs.unit_cost_growth", -0.9, 5, "the growth in unit cost (a fraction)"),
    ("costs.fixed_costs", 0, BILLION, "fixed costs"), ("costs.fixed_cost_growth", -0.9, 5, "the growth in fixed costs (a fraction)"),
    ("equipment.initial", 0, BILLION, "equipment bought at the start"), ("equipment.yearly", 0, BILLION, "equipment bought each year"),
    ("equipment.depreciation_rate", 1e-9, 1, "the depreciation rate (a fraction of the value, above 0, at most 1)"),
    ("working_capital.receivable_days", 0, 365, "receivable days"), ("working_capital.inventory_days", 0, 365, "stock days"),
    ("working_capital.payable_days", 0, 365, "payable days"),
    ("funding.equity", 0, BILLION, "the owners' money"), ("funding.loan", 0, BILLION, "the loan"),
    ("funding.interest_rate", 0, 0.999, "the interest rate (a fraction)"), ("funding.loan_years", 1, 50, "the years to repay the loan"),
    ("tax_rate", 0, 0.999, "the tax rate (a fraction)"),
]
EXAMPLE = {
    "name": "Harbour Bike Repairs",
    "description": "A one-workshop bike repair and parts business on a harbour road: repairs and servicing, with parts sold at the counter.",
    "currency": "£", "first_year": 2027, "years": 3, "illustrative": True,
    "sales": {"units": 2400, "unit_growth": 0.12, "price": 45, "price_growth": 0.03},
    "costs": {"unit_cost": 18, "unit_cost_growth": 0.02, "fixed_costs": 48000, "fixed_cost_growth": 0.03},
    "equipment": {"initial": 30000, "yearly": 4000, "depreciation_rate": 0.2},
    "working_capital": {"receivable_days": 15, "inventory_days": 45, "payable_days": 30},
    "funding": {"equity": 25000, "loan": 20000, "interest_rate": 0.07, "loan_years": 5},
    "tax_rate": 0.19,
}


# emoji and icon characters (model_check.py's K11 uses the same set): the pictograph and dingbat blocks, the emoji
# punctuation and letter-like signs (‼ ⁉ ℹ Ⓜ), the arrows Unicode lists as emoji (↔ … ↙, ↩ ↪; → is text), and the
# emoji presentation selector. © ® ™ are text.
ICONS = re.compile("[\u203c\u2049\u2139\u2194-\u2199\u21a9\u21aa\u24c2\u2295-\u22a1\u2300-\u23ff\u25a0-\u25ff"
                   "\u2600-\u27bf\u2b00-\u2bff\u3030\u303d\u3297\u3299\U0001f000-\U0001faff\ufe0f\u20e3]")
CONTROL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ud800-\udfff]")   # not writable into the workbook's XML


def safe_number(v):
    """a number as a refusal can print it, however large"""
    try:
        return f"{float(v):g}"
    except OverflowError:
        return f"a number of {len(str(abs(v)))} digits"


def get(b, path):
    for k in path.split("."):
        b = b[k]
    return b


def _finite(v):
    try:
        return math.isfinite(v)
    except OverflowError:                 # an integer too large to be a float: not a number this model can use
        return False


def _no_repeats(pairs):
    seen = {}
    for k, v in pairs:
        if k in seen:
            raise ValueError(f"the key '{k}' appears twice in one object")
        seen[k] = v
    return seen


def loads(text):
    try:
        return json.loads(text, object_pairs_hook=_no_repeats), None
    except ValueError as e:
        return None, f"not valid JSON: {e}"


def problems(b):
    """[(code, message)] — empty when the business can be modelled"""
    out = []

    def typed(where, obj, spec):
        if not isinstance(obj, dict):
            out.append(("B01", f"{where} should be an object")); return
        for k in obj:
            if k not in spec:
                out.append(("B01", f"{where} has a field '{k}' this skill does not know (known: {', '.join(spec)})"))
        for k, kind in spec.items():
            name = f"{where}.{k}" if where != "the business" else k
            if k not in obj:
                if not (isinstance(kind, str) and kind.endswith("?")):
                    out.append(("B01", f"{name} is missing"))
                continue
            v = obj[k]
            if isinstance(kind, dict):
                typed(name, v, kind)
            elif kind.startswith("str"):
                if not isinstance(v, str) or not v.strip():
                    out.append(("B01", f"{name} should be text"))
                elif CONTROL.search(v):
                    out.append(("B01", f"{name} has a control character or a broken one (U+{ord(CONTROL.search(v).group(0)):04X}) that the workbook cannot hold"))
                elif ICONS.search(v):
                    out.append(("B01", f"{name} has an emoji or icon character ({ICONS.search(v).group(0)}); the workbook is text and numbers"))
            elif kind == "bool":
                if not isinstance(v, bool):
                    out.append(("B01", f"{name} should be true or false"))
            elif kind == "int":
                if isinstance(v, bool) or not isinstance(v, int):
                    out.append(("B01", f"{name} should be a whole number, not {json.dumps(v)}"))
            elif isinstance(v, int) and not isinstance(v, bool) and not _finite(v):
                out.append(("B01", f"{name} is a number of {len(str(abs(v)))} digits, too large for any model"))
            elif isinstance(v, bool) or not isinstance(v, (int, float)) or not _finite(v):
                out.append(("B01", f"{name} should be a number, not {json.dumps(v, ensure_ascii=False)[:40]}"))

    typed("the business", b, SPEC)
    if out:
        return out
    for path, lo, hi, words in RANGES:
        v = get(b, path)
        if v < lo or v > hi:
            out.append(("B02", f"{path} = {safe_number(v)}: {words} should be between {lo:g} and {hi:g}"))
    if len(b["currency"]) > 3:
        out.append(("B02", f"currency '{b['currency']}' should be a symbol or a three-letter code"))
    if b["illustrative"] and b.get("sources"):
        out.append(("B03", "the numbers are marked illustrative and also given sources; say which: made up (illustrative) or real (sources)"))
    if not b["illustrative"] and not b.get("sources"):
        out.append(("B03", "the numbers are not marked illustrative, so 'sources' has to say where they come from"))
    return out


# ---- the layout: where everything goes. model_check.py reads it from here, so the two cannot drift apart ---------------
ASSUMPTIONS = [   # (row, label, path, format, note)
    (6, "Units sold in the first year", "sales.units", "count", "units"),
    (7, "Growth in units sold each year", "sales.unit_growth", "pct", ""),
    (8, "Price of one unit in the first year", "sales.price", "money2", ""),
    (9, "Growth in price each year", "sales.price_growth", "pct", ""),
    (11, "Cost of one unit in the first year", "costs.unit_cost", "money2", ""),
    (12, "Growth in unit cost each year", "costs.unit_cost_growth", "pct", ""),
    (13, "Fixed operating costs in the first year", "costs.fixed_costs", "money", ""),
    (14, "Growth in fixed costs each year", "costs.fixed_cost_growth", "pct", ""),
    (16, "Equipment bought at the start", "equipment.initial", "money", ""),
    (17, "Equipment bought each year", "equipment.yearly", "money", ""),
    (18, "Depreciation: share of the equipment's value at the start of each year", "equipment.depreciation_rate", "pct", "reducing balance"),
    (20, "Days customers take to pay", "working_capital.receivable_days", "days", "days"),
    (21, "Days of stock held", "working_capital.inventory_days", "days", "days"),
    (22, "Days taken to pay suppliers", "working_capital.payable_days", "days", "days"),
    (24, "Owners' money put in at the start", "funding.equity", "money", ""),
    (25, "Loan taken at the start", "funding.loan", "money", ""),
    (26, "Interest rate on the loan", "funding.interest_rate", "pct", "on the balance at the start of each year"),
    (27, "Years to repay the loan", "funding.loan_years", "count", "equal yearly repayments"),
    (29, "Tax rate on profit", "tax_rate", "pct", "charged only on a profit"),
]
SECTIONS = {5: "Sales", 10: "Costs", 15: "Equipment", 19: "Working capital", 23: "Funding", 28: "Tax"}
A = {path: row for row, _, path, _, _ in ASSUMPTIONS}          # assumption path -> row on the Assumptions sheet


def aref(path):
    return f"Assumptions!$B${A[path]}"


IS = {"units": 4, "price": 5, "revenue": 6, "ucost": 7, "cogs": 8, "gross": 9, "fixed": 10, "dep": 11, "ebit": 12,
      "interest": 13, "pbt": 14, "tax": 15, "net": 16}
IS_LABELS = {"units": "Units sold", "price": "Price of one unit", "revenue": "Revenue", "ucost": "Cost of one unit",
             "cogs": "Cost of sales", "gross": "Gross profit", "fixed": "Fixed operating costs", "dep": "Depreciation",
             "ebit": "Operating profit", "interest": "Interest on the loan", "pbt": "Profit before tax", "tax": "Tax",
             "net": "Net profit"}
BS = {"cash": 4, "recv": 5, "stock": 6, "equip": 7, "assets": 8, "pay": 10, "loan": 11, "liab": 12, "capital": 14,
      "retained": 15, "equity": 16, "le": 18, "diff": 19}
BS_LABELS = {"cash": "Cash", "recv": "Money owed by customers", "stock": "Stock", "equip": "Equipment",
             "assets": "Total assets", "pay": "Money owed to suppliers", "loan": "Loan", "liab": "Total liabilities",
             "capital": "Owners' money", "retained": "Profit kept in the business", "equity": "Total equity",
             "le": "Total liabilities and equity", "diff": "Assets minus liabilities and equity"}
CF = {"net": 4, "dep": 5, "recv": 6, "stock": 7, "pay": 8, "ops": 9, "capex": 10, "inv": 11, "repay": 12, "fin": 13,
      "change": 14, "open": 15, "close": 16}
CF_LABELS = {"net": "Net profit", "dep": "Depreciation added back", "recv": "Change in money owed by customers",
             "stock": "Change in stock", "pay": "Change in money owed to suppliers", "ops": "Cash from operations",
             "capex": "Equipment bought", "inv": "Cash from investing", "repay": "Loan repaid", "fin": "Cash from financing",
             "change": "Change in cash", "open": "Cash at the start of the year", "close": "Cash at the end of the year"}
CK = {"balance": 4, "cash": 5, "retained": 6, "equip": 7, "loan": 8, "result": 10, "negative": 11}
CK_LABELS = {"balance": "Assets minus liabilities and equity", "cash": "Cash on the balance sheet minus cash from the cash flow",
             "retained": "Profit kept: this year minus (last year plus net profit)",
             "equip": "Equipment: this year minus (last year plus bought minus depreciation)",
             "loan": "Loan: this year minus (last year minus repaid)", "result": "Result",
             "negative": "Cash below zero at the start or at a year end"}
SHEETS = ["Assumptions", "Income statement", "Balance sheet", "Cash flow", "Checks", "Charts"]
TOTALS_IS, DOUBLE_IS = {"gross", "ebit", "pbt"}, {"net"}
TOTALS_BS, DOUBLE_BS = {"assets", "liab", "equity"}, {"le"}
TOTALS_CF, DOUBLE_CF = {"ops", "inv", "fin", "change"}, {"close"}


def ycol(i):
    """the column of year i (1-based) on the income statement and the cash flow; the balance sheet has Start in B, so
    year i sits one column to the right there"""
    return 1 + i


def bcol(i):
    return 2 + i                      # i = 0 is the start


def zero_within_penny(expr):
    """a check's difference, as an exact 0 when it is within half a penny: floating point leaves -7.3e-12 behind, and a
    viewer that ignores number formats (Quick Look) would print it as -7.27596E-12"""
    return f"IF(ABS({expr})<0.005,0,{expr})"


def formulas(years):
    """{(sheet, row, col): formula} for everything on the statements and the checks — the model, as formulas"""
    f = {}
    ISn, BSn, CFn = "'Income statement'", "'Balance sheet'", "'Cash flow'"
    for i in range(1, years + 1):
        c, p = O.col_letter(ycol(i)), O.col_letter(ycol(i - 1))           # this year / last year on IS and CF
        bc, bp = O.col_letter(bcol(i)), O.col_letter(bcol(i - 1))         # this year / last year on the balance sheet
        first = i == 1
        s = "Income statement"
        f[(s, IS["units"], ycol(i))] = aref("sales.units") if first else f"{p}{IS['units']}*(1+{aref('sales.unit_growth')})"
        f[(s, IS["price"], ycol(i))] = aref("sales.price") if first else f"{p}{IS['price']}*(1+{aref('sales.price_growth')})"
        f[(s, IS["revenue"], ycol(i))] = f"{c}{IS['units']}*{c}{IS['price']}"
        f[(s, IS["ucost"], ycol(i))] = aref("costs.unit_cost") if first else f"{p}{IS['ucost']}*(1+{aref('costs.unit_cost_growth')})"
        f[(s, IS["cogs"], ycol(i))] = f"{c}{IS['units']}*{c}{IS['ucost']}"
        f[(s, IS["gross"], ycol(i))] = f"{c}{IS['revenue']}-{c}{IS['cogs']}"
        f[(s, IS["fixed"], ycol(i))] = aref("costs.fixed_costs") if first else f"{p}{IS['fixed']}*(1+{aref('costs.fixed_cost_growth')})"
        f[(s, IS["dep"], ycol(i))] = f"{BSn}!{bp}{BS['equip']}*{aref('equipment.depreciation_rate')}"
        f[(s, IS["ebit"], ycol(i))] = f"{c}{IS['gross']}-{c}{IS['fixed']}-{c}{IS['dep']}"
        f[(s, IS["interest"], ycol(i))] = f"{BSn}!{bp}{BS['loan']}*{aref('funding.interest_rate')}"
        f[(s, IS["pbt"], ycol(i))] = f"{c}{IS['ebit']}-{c}{IS['interest']}"
        f[(s, IS["tax"], ycol(i))] = f"MAX(0,{c}{IS['pbt']})*{aref('tax_rate')}"
        f[(s, IS["net"], ycol(i))] = f"{c}{IS['pbt']}-{c}{IS['tax']}"
        s = "Cash flow"
        f[(s, CF["net"], ycol(i))] = f"{ISn}!{c}{IS['net']}"
        f[(s, CF["dep"], ycol(i))] = f"{ISn}!{c}{IS['dep']}"
        f[(s, CF["recv"], ycol(i))] = f"{BSn}!{bp}{BS['recv']}-{BSn}!{bc}{BS['recv']}"
        f[(s, CF["stock"], ycol(i))] = f"{BSn}!{bp}{BS['stock']}-{BSn}!{bc}{BS['stock']}"
        f[(s, CF["pay"], ycol(i))] = f"{BSn}!{bc}{BS['pay']}-{BSn}!{bp}{BS['pay']}"
        f[(s, CF["ops"], ycol(i))] = f"SUM({c}{CF['net']}:{c}{CF['pay']})"
        f[(s, CF["capex"], ycol(i))] = f"-{aref('equipment.yearly')}"
        f[(s, CF["inv"], ycol(i))] = f"{c}{CF['capex']}"
        f[(s, CF["repay"], ycol(i))] = f"-MIN({BSn}!{bp}{BS['loan']},{aref('funding.loan')}/{aref('funding.loan_years')})"
        f[(s, CF["fin"], ycol(i))] = f"{c}{CF['repay']}"
        f[(s, CF["change"], ycol(i))] = f"{c}{CF['ops']}+{c}{CF['inv']}+{c}{CF['fin']}"
        f[(s, CF["open"], ycol(i))] = f"{BSn}!{bp}{BS['cash']}"
        f[(s, CF["close"], ycol(i))] = f"{c}{CF['open']}+{c}{CF['change']}"
        s = "Balance sheet"
        f[(s, BS["cash"], bcol(i))] = f"{CFn}!{c}{CF['close']}"
        f[(s, BS["recv"], bcol(i))] = f"{ISn}!{c}{IS['revenue']}*{aref('working_capital.receivable_days')}/365"
        f[(s, BS["stock"], bcol(i))] = f"{ISn}!{c}{IS['cogs']}*{aref('working_capital.inventory_days')}/365"
        f[(s, BS["equip"], bcol(i))] = f"{bp}{BS['equip']}+{aref('equipment.yearly')}-{ISn}!{c}{IS['dep']}"
        f[(s, BS["pay"], bcol(i))] = f"{ISn}!{c}{IS['cogs']}*{aref('working_capital.payable_days')}/365"
        f[(s, BS["loan"], bcol(i))] = f"{bp}{BS['loan']}+{CFn}!{c}{CF['repay']}"
        f[(s, BS["capital"], bcol(i))] = f"{bp}{BS['capital']}"
        f[(s, BS["retained"], bcol(i))] = f"{bp}{BS['retained']}+{ISn}!{c}{IS['net']}"
    for i in range(0, years + 1):
        bc = O.col_letter(bcol(i))
        s = "Balance sheet"
        if i == 0:
            # the owners' money and the loan less the equipment bought; within half a penny of nothing it is nothing, or
            # floating point leaves -3.6e-12 behind when they cancel exactly (17,129.37 + 10,590.90 - 27,720.27), and
            # the Checks sheet would say the model needs money it does not include
            f[(s, BS["cash"], bcol(0))] = zero_within_penny(f"{aref('funding.equity')}+{aref('funding.loan')}-{aref('equipment.initial')}")
            f[(s, BS["equip"], bcol(0))] = aref("equipment.initial")
            f[(s, BS["loan"], bcol(0))] = aref("funding.loan")
            f[(s, BS["capital"], bcol(0))] = aref("funding.equity")
        f[(s, BS["assets"], bcol(i))] = f"SUM({bc}{BS['cash']}:{bc}{BS['equip']})"
        f[(s, BS["liab"], bcol(i))] = f"SUM({bc}{BS['pay']}:{bc}{BS['loan']})"
        f[(s, BS["equity"], bcol(i))] = f"SUM({bc}{BS['capital']}:{bc}{BS['retained']})"
        f[(s, BS["le"], bcol(i))] = f"{bc}{BS['liab']}+{bc}{BS['equity']}"
        f[(s, BS["diff"], bcol(i))] = zero_within_penny(f"{bc}{BS['assets']}-{bc}{BS['le']}")
        s = "Checks"
        f[(s, CK["balance"], bcol(i))] = f"{BSn}!{bc}{BS['diff']}"
        if i:
            bp, c = O.col_letter(bcol(i - 1)), O.col_letter(ycol(i))
            f[(s, CK["cash"], bcol(i))] = zero_within_penny(f"{BSn}!{bc}{BS['cash']}-{CFn}!{c}{CF['close']}")
            f[(s, CK["retained"], bcol(i))] = zero_within_penny(f"{BSn}!{bc}{BS['retained']}-({BSn}!{bp}{BS['retained']}+{ISn}!{c}{IS['net']})")
            f[(s, CK["equip"], bcol(i))] = zero_within_penny(f"{BSn}!{bc}{BS['equip']}-({BSn}!{bp}{BS['equip']}+{aref('equipment.yearly')}-{ISn}!{c}{IS['dep']})")
            f[(s, CK["loan"], bcol(i))] = zero_within_penny(f"{BSn}!{bc}{BS['loan']}-({BSn}!{bp}{BS['loan']}+{CFn}!{c}{CF['repay']})")
    first, last = O.col_letter(bcol(0)), O.col_letter(bcol(years))
    span = f"{first}{CK['balance']}:{last}{CK['loan']}"
    f[("Checks", CK["result"], 2)] = (f'IF(AND(MAX({span})<0.005,MIN({span})>-0.005),'
                                      f'"Balanced: every check is zero to the penny","NOT balanced: see the rows above")')
    cash = f"{BSn}!{first}{BS['cash']}:{last}{BS['cash']}"
    # short only by more than half a penny: a year-end cash of exactly nothing can be stored as -3.6e-12, and it is not
    # rounded (a year-end cash can be fractions of a penny, and they carry into the next year)
    f[("Checks", CK["negative"], 2)] = f'IF(MIN({cash})<-0.005,"Yes: the model needs money it does not include","No")'
    return f


def compute(book_formulas, inputs):
    """every formula's value, worked out in dependency order with the same evaluator model_check.py uses"""
    values = dict(inputs)                       # (sheet, "B6") -> value
    trees = {(s, O.ref(r, c)): X.parse(fx) for (s, r, c), fx in book_formulas.items()}
    pending = set(trees)
    while pending:
        progressed = False
        for key in sorted(pending):
            s, cell = key
            deps = [d for d in X.refs(trees[key], s) if d in trees]
            if all(d in values for d in deps):
                values[key] = X.evaluate(trees[key], s, lambda sh, c: values.get((sh, c)))
                pending.discard(key)
                progressed = True
        if not progressed:
            raise RuntimeError("the formulas refer to each other in a circle: " + ", ".join(f"{s}!{c}" for s, c in sorted(pending)[:5]))
    return values


def direct(b):
    """the same model worked out again in plain Python, from the business, without a single formula: statements by year"""
    n = b["years"]
    s, co, eq, wc, fu = b["sales"], b["costs"], b["equipment"], b["working_capital"], b["funding"]
    out = {"start": {}, "years": []}
    cash0 = fu["equity"] + fu["loan"] - eq["initial"]
    prev = {"cash": cash0, "recv": 0.0, "stock": 0.0, "equip": float(eq["initial"]), "pay": 0.0, "loan": float(fu["loan"]),
            "capital": float(fu["equity"]), "retained": 0.0}
    out["start"] = dict(prev)
    units, price, ucost, fixed = s["units"], s["price"], co["unit_cost"], co["fixed_costs"]
    for i in range(n):
        if i:
            units, price = units * (1 + s["unit_growth"]), price * (1 + s["price_growth"])
            ucost, fixed = ucost * (1 + co["unit_cost_growth"]), fixed * (1 + co["fixed_cost_growth"])
        revenue, cogs = units * price, units * ucost
        dep = prev["equip"] * eq["depreciation_rate"]
        interest = prev["loan"] * fu["interest_rate"]
        pbt = revenue - cogs - fixed - dep - interest
        tax = max(0.0, pbt) * b["tax_rate"]
        net = pbt - tax
        recv, stock, pay = revenue * wc["receivable_days"] / 365, cogs * wc["inventory_days"] / 365, cogs * wc["payable_days"] / 365
        repay = min(prev["loan"], fu["loan"] / fu["loan_years"])
        ops = net + dep - (recv - prev["recv"]) - (stock - prev["stock"]) + (pay - prev["pay"])
        cash = prev["cash"] + ops - eq["yearly"] - repay
        now = {"cash": cash, "recv": recv, "stock": stock, "equip": prev["equip"] + eq["yearly"] - dep, "pay": pay,
               "loan": prev["loan"] - repay, "capital": prev["capital"], "retained": prev["retained"] + net}
        out["years"].append({"revenue": revenue, "cogs": cogs, "dep": dep, "interest": interest, "tax": tax, "net": net,
                             "ops": ops, "repay": repay, **now})
        prev = now
    return out


def compare_direct(values, b):
    """[(where, workbook value, python value)] for every statement number the two workings disagree on"""
    d = direct(b)
    wrong = []

    def near(x, y):
        return x is not None and abs(x - y) < 0.005

    for key, name in (("cash", "cash"), ("equip", "equip"), ("loan", "loan"), ("capital", "capital")):
        v = values.get(("Balance sheet", O.ref(BS[name], bcol(0))))
        if not near(v, d["start"][key]):
            wrong.append((f"Balance sheet!{O.ref(BS[name], bcol(0))}", v, d["start"][key]))
    for i, y in enumerate(d["years"], 1):
        pairs = [("Income statement", IS["revenue"], ycol(i), y["revenue"]), ("Income statement", IS["cogs"], ycol(i), y["cogs"]),
                 ("Income statement", IS["dep"], ycol(i), y["dep"]), ("Income statement", IS["interest"], ycol(i), y["interest"]),
                 ("Income statement", IS["tax"], ycol(i), y["tax"]), ("Income statement", IS["net"], ycol(i), y["net"]),
                 ("Cash flow", CF["ops"], ycol(i), y["ops"]), ("Cash flow", CF["repay"], ycol(i), -y["repay"]),
                 ("Cash flow", CF["close"], ycol(i), y["cash"])]
        pairs += [("Balance sheet", BS[k], bcol(i), y[k]) for k in ("cash", "recv", "stock", "equip", "pay", "loan", "capital", "retained")]
        for sheet, row, col, want in pairs:
            v = values.get((sheet, O.ref(row, col)))
            if not near(v, want):
                wrong.append((f"{sheet}!{O.ref(row, col)}", v, want))
    return wrong


# ---- the workbook ----------------------------------------------------------------------------------------------------
PALETTE = {"band": "053333", "on_band": "F7F0E9", "ink": "1D1B24", "muted": "6E6674", "input": "F8E4E7", "good": "1F6B4F",
           "bad": "A33B2F", "rule": "4F4A57", "revenue": "3CA77A", "profit": "CA6980", "cash": "053333"}


def styles(wb, currency):
    s = wb.styles
    money = '#,##0;(#,##0);"–"'
    s.add("title", font={"name": "Fraunces", "size": 16, "bold": True})
    s.add("note", font={"size": 9, "italic": True, "color": PALETTE["muted"]})
    s.add("head", font={"bold": True, "color": PALETTE["on_band"]}, fill=PALETTE["band"])
    s.add("headnum", font={"bold": True, "color": PALETTE["on_band"]}, fill=PALETTE["band"], align="right")
    s.add("section", font={"bold": True, "size": 11})
    s.add("label", indent=1)
    s.add("labeltotal", font={"bold": True})
    s.add("money", font={"name": "Space Mono"}, numfmt=money)
    s.add("moneytotal", font={"name": "Space Mono", "bold": True}, numfmt=money, border=("thin", False))
    s.add("moneydouble", font={"name": "Space Mono", "bold": True}, numfmt=money, border=("thin", True))
    s.add("money2", font={"name": "Space Mono"}, numfmt='#,##0.00;(#,##0.00);"–"')
    s.add("count", font={"name": "Space Mono"}, numfmt="#,##0")
    # a difference within half a penny reads 0.00, never -0.00; a real one keeps its sign
    s.add("check", font={"name": "Space Mono"}, numfmt="0.00;-0.00;0.00")
    s.add("in_money", font={"name": "Space Mono"}, numfmt=money, fill=PALETTE["input"])
    s.add("in_money2", font={"name": "Space Mono"}, numfmt="#,##0.00", fill=PALETTE["input"])
    s.add("in_pct", font={"name": "Space Mono"}, numfmt="0.0%", fill=PALETTE["input"])
    s.add("in_count", font={"name": "Space Mono"}, numfmt="#,##0", fill=PALETTE["input"])
    s.add("in_days", font={"name": "Space Mono"}, numfmt="0", fill=PALETTE["input"])
    s.add("verdict", font={"bold": True, "size": 12, "color": PALETTE["ink"]})      # plain ink: it reads right whatever it says
    s.add("negline", font={"bold": True})
    s.add("text", font={"color": PALETTE["ink"]})


def header(ws, title, note, heads, width_label=36, width_num=12):
    ws.set(1, 1, title, style="title")
    ws.set(2, 1, note, style="note")
    ws.width(1, width_label)
    for j, h in enumerate(heads):
        ws.set(3, 1 + j, h, style="head" if j == 0 else "headnum")
        if j:
            ws.width(1 + j, width_num)
    ws.freeze = "B4"


def build(b, bend=None):
    """the workbook for a business, every formula's value filled in; returns (workbook, values). bend(formulas) may
    change the formulas first — model_check.py's self-test uses it to make models that are wrong in one known way"""
    n, cur = b["years"], b["currency"]
    years = [str(b["first_year"] + i) for i in range(n)]
    wb = O.Workbook(title=f"{b['name']}: three-statement model", creator="nk-model")
    styles(wb, cur)
    sh = {name: wb.sheet(name) for name in SHEETS}
    # Assumptions
    a = sh["Assumptions"]
    a.width(1, 62); a.width(2, 14); a.width(3, 40)
    a.set(1, 1, f"Assumptions: {b['name']}", style="title")
    a.set(2, 1, b["description"], style="note")
    a.set(3, 1, ("Illustrative: the numbers are made up for a demonstration. " if b["illustrative"] else f"Where the numbers come from: {b['sources']}. ") +
             "This is a model of how the numbers fit together, not financial advice.", style="note")
    for row, title in SECTIONS.items():
        a.set(row, 1, title, style="section")
    inputs = {}
    fmt_style = {"money": "in_money", "money2": "in_money2", "pct": "in_pct", "count": "in_count", "days": "in_days"}
    for row, label, path, fmt, note in ASSUMPTIONS:
        v = get(b, path)
        a.set(row, 1, label, style="label"); a.set(row, 2, v, style=fmt_style[fmt])
        a.set(row, 3, "; ".join(x for x in ((cur if fmt in ("money", "money2") else ""), note) if x), style="note")
        inputs[("Assumptions", f"B{row}")] = float(v)
    a.set(31, 1, "How the model works", style="section")
    how = [f"Years run from {years[0]} to {years[-1]}; the balance sheet also shows the start, before the first year.",
           "Revenue is units sold times price; cost of sales is units sold times the cost of one unit.",
           "Depreciation is a fixed share of the equipment's value at the start of each year (reducing balance).",
           "Money owed by customers, stock and money owed to suppliers are the year's revenue or cost of sales times days, divided by 365.",
           "Interest is charged on the loan balance at the start of each year; the loan is repaid in equal yearly amounts.",
           "Tax is charged only on a profit; a loss is not carried forward. Nothing is paid out to the owners.",
           "The cash flow starts from net profit and adjusts it for depreciation and for money tied up in customers, stock and suppliers.",
           "Totals are worked out before rounding, so a total can be a few units off the lines it shows; prices, unit costs and checks show pennies.",
           "Units sold grow by a percentage each year, so after the first year they are not whole numbers; revenue is worked out from the exact units.",
           "The pink cells are the inputs. Every number the other sheets show is a formula that points back here."]
    for k, line in enumerate(how):
        a.set(32 + k, 1, line, style="text")
    a.freeze = "A4"
    # statements: labels and headers
    ist, bst, cft, ck = sh["Income statement"], sh["Balance sheet"], sh["Cash flow"], sh["Checks"]
    header(ist, f"Income statement: {b['name']}", f"{cur}, for each year, rounded (see the Assumptions sheet)", ["Year"] + years)
    header(bst, f"Balance sheet: {b['name']}", f"{cur}, at the end of each year, rounded; Start is before the first year", ["At the end of", "Start"] + years)
    header(cft, f"Cash flow: {b['name']}", f"{cur}, for each year, rounded; money in is positive, money out in brackets", ["Year"] + years)
    header(ck, f"Checks: {b['name']}", "Each line should be zero in every year. The model is only handed over when they all are.", ["Check", "Start"] + years, width_label=58)
    for key, row in IS.items():
        ist.set(row, 1, IS_LABELS[key], style="labeltotal" if key in TOTALS_IS | DOUBLE_IS else "label")
    for key, row in BS.items():
        bst.set(row, 1, BS_LABELS[key], style="labeltotal" if key in TOTALS_BS | DOUBLE_BS or key == "diff" else "label")
    for key, row in CF.items():
        cft.set(row, 1, CF_LABELS[key], style="labeltotal" if key in TOTALS_CF | DOUBLE_CF else "label")
    for key, row in CK.items():
        ck.set(row, 1, CK_LABELS[key], style="labeltotal" if key in ("result", "negative") else "label")
    # formulas, values, and the style each cell takes
    fx = formulas(n)
    if bend:
        bend(fx)
    values = compute(fx, inputs)

    def style_for(sheet, row):
        key = {s: {v: k for k, v in m.items()} for s, m in (("Income statement", IS), ("Balance sheet", BS), ("Cash flow", CF), ("Checks", CK))}[sheet].get(row)
        if sheet == "Checks":            # the verdict and the cash line in plain bold ink: a colour would read right whatever they say
            return "verdict" if key == "result" else "negline" if key == "negative" else "check"
        if sheet == "Income statement" and key in ("units",):
            return "count"
        if sheet == "Income statement" and key in ("price", "ucost"):
            return "money2"
        if sheet == "Balance sheet" and key == "diff":
            return "check"
        totals = {"Income statement": (TOTALS_IS, DOUBLE_IS), "Balance sheet": (TOTALS_BS, DOUBLE_BS), "Cash flow": (TOTALS_CF, DOUBLE_CF)}[sheet]
        return "moneydouble" if key in totals[1] else "moneytotal" if key in totals[0] else "money"

    for (s, r, c), f in fx.items():
        sh[s].set(r, c, values[(s, O.ref(r, c))], formula=f, style=style_for(s, r))
    # charts: what the business earns, and its cash
    ch = sh["Charts"]
    ch.set(1, 1, f"Charts: {b['name']}", style="title")
    ch.set(2, 1, f"{cur}. The charts read the statements; each one also carries a copy of the numbers it draws.", style="note")
    ch.width(1, 12)
    first, last = O.col_letter(ycol(1)), O.col_letter(ycol(n))
    cats = (f"'Income statement'!${first}$3:${last}$3", years)
    ch.chart(kind="bar", title="Revenue and net profit", at=(1, 4, 10, 22), number_format="#,##0", categories=cats,
             series=[("Revenue", f"'Income statement'!${first}${IS['revenue']}:${last}${IS['revenue']}",
                      [values[("Income statement", O.ref(IS["revenue"], ycol(i)))] for i in range(1, n + 1)], PALETTE["revenue"]),
                     ("Net profit", f"'Income statement'!${first}${IS['net']}:${last}${IS['net']}",
                      [values[("Income statement", O.ref(IS["net"], ycol(i)))] for i in range(1, n + 1)], PALETTE["profit"])])
    ch.chart(kind="line", title="Cash at the end of each year", at=(1, 24, 10, 42), number_format="#,##0", categories=cats,
             series=[("Cash", f"'Cash flow'!${first}${CF['close']}:${last}${CF['close']}",
                      [values[("Cash flow", O.ref(CF["close"], ycol(i)))] for i in range(1, n + 1)], PALETTE["cash"])])
    return wb, values


def verdict(values, b):
    """(balanced?, [problems]) from the computed values: every check zero to the penny, and the second working agrees"""
    bad = []
    n = b["years"]
    for key in ("balance", "cash", "retained", "equip", "loan"):
        for i in range(0 if key == "balance" else 1, n + 1):
            v = values.get(("Checks", O.ref(CK[key], bcol(i))))
            if v is None or abs(v) >= 0.005:
                bad.append(f"Checks!{O.ref(CK[key], bcol(i))} ({CK_LABELS[key]}) is {v!r}, not zero")
    for where, got, want in compare_direct(values, b):
        bad.append(f"{where} is {got!r} in the workbook but {want!r} worked out directly")
    return not bad, bad


def make(src, out):
    try:
        text = io.open(src, encoding="utf-8").read()
    except OSError as e:
        print(f"cannot read {src}: {e}"); return 2
    b, why = loads(text)
    if why:
        print("B01 " + why); print("nothing written"); return 1
    found = problems(b)
    if found:
        for code, m in found:
            print(f"{code} {m}")
        print(f"{len(found)} problem(s); nothing written"); return 1
    wb, values = build(b)
    ok, bad = verdict(values, b)
    if not ok:
        for m in bad[:12]:
            print("UNBALANCED " + m)
        print(f"the model does not balance or tie ({len(bad)} place(s)); nothing written"); return 1
    size = wb.save(out)
    n = b["years"]
    cash = [values[("Cash flow", O.ref(CF["close"], ycol(i)))] for i in range(1, n + 1)]
    start_cash = values[("Balance sheet", O.ref(BS["cash"], bcol(0)))]
    net = [values[("Income statement", O.ref(IS["net"], ycol(i)))] for i in range(1, n + 1)]
    print(f"wrote {out}: {b['name']}, {b['first_year']}-{b['first_year'] + n - 1} · {size:,} bytes · balanced to the penny in all {n + 1} balance sheets")
    # display only: a cash of exactly nothing can be stored as -3.6e-12, which would print as "-0"
    whole = lambda v: "0" if f"{v:,.0f}" == "-0" else f"{v:,.0f}"
    print(f"net profit {', '.join(whole(v) for v in net)} · cash at the end of each year {', '.join(whole(v) for v in cash)} ({b['currency']})")
    # below zero by more than half a penny, as the Checks sheet judges it
    low = [("Start", start_cash)] * (start_cash < -0.005) + [(str(b["first_year"] + i), v) for i, v in enumerate(cash) if v < -0.005]
    if low:
        print(f"note: cash is below zero in {', '.join(y for y, _ in low)}: the model needs money it does not include")
    print(f"next: python3 {os.path.join(HERE, 'model_check.py')} {out}")
    return 0


def selftest():
    import tempfile, subprocess
    ok = []

    def say(cond, text):
        ok.append(bool(cond)); print(f"  {'PASS' if cond else 'FAIL'}  {text}")

    say(not problems(EXAMPLE), "the example business meets every rule")
    wb, values = build(EXAMPLE)
    good, bad = verdict(values, EXAMPLE)
    say(good, "the example balances to the penny and the direct working agrees" + "".join("\n        " + x for x in bad[:4]))
    n = EXAMPLE["years"]
    for i in range(n + 1):
        a = values[("Balance sheet", O.ref(BS["assets"], bcol(i)))]
        le = values[("Balance sheet", O.ref(BS["le"], bcol(i)))]
        if abs(a - le) >= 0.005:
            say(False, f"balance sheet {i} does not balance")
    say(values[("Checks", "B10")].startswith("Balanced"), "the checks sheet says Balanced")
    # the Assumptions sheet's first column is as wide as its longest line, and a preview pushes the inputs beside it out
    # of view when it grows: in a 1000-px preview a 138-character line keeps them inside (right edge 992 px) and a
    # 140-character one does not (1,004 px), measured 2026-09-23. Characters stand in for the width a line takes.
    widest = 138
    first = next(sh for sh in wb.sheets if sh.name == "Assumptions")
    long = [f"{O.ref(r, c)} ({len(v)})" for (r, c), (kind, v, _, _) in sorted(first.cells.items())
            if c == 1 and r > 2 and kind == "s" and len(v) > widest]      # A1 and A2 hold the business's own name and words
    say(not long, f"every line this skill writes in the Assumptions sheet's first column is at most {widest} characters"
                  + (f": {', '.join(long)}" if long else ""))
    for bizname, tweak in (("a business that loses money every year", {"costs": dict(EXAMPLE["costs"], fixed_costs=200000)}),
                           ("a business with no loan", {"funding": dict(EXAMPLE["funding"], loan=0)}),
                           ("ten years and a loan repaid in two", {"years": 10, "funding": dict(EXAMPLE["funding"], loan_years=2)}),
                           ("no stock, instant payment", {"working_capital": {"receivable_days": 0, "inventory_days": 0, "payable_days": 0}})):
        b = json.loads(json.dumps(EXAMPLE)); b.update(tweak)
        _, v = build(b)
        g, bad = verdict(v, b)
        say(g and not problems(b), f"{bizname}: balanced and agreed" + "".join("\n        " + x for x in bad[:3]))
    # a formula bent in the builder is caught before anything is written
    def bent(f):
        f[("Balance sheet", BS["recv"], bcol(1))] += "*1.01"
    _, v = build(EXAMPLE, bend=bent)
    g, bad = verdict(v, EXAMPLE)
    say(not g and any("Money owed" in x or "Balance sheet!C5" in x or "Checks" in x for x in bad), "a formula bent by one percent is caught: the model no longer balances or agrees")
    with tempfile.TemporaryDirectory() as t:
        src, out = os.path.join(t, "b.json"), os.path.join(t, "m.xlsx")
        io.open(src, "w", encoding="utf-8").write(json.dumps(EXAMPLE))
        r = subprocess.run([sys.executable, __file__, src, "-o", out], capture_output=True, text=True)
        say(r.returncode == 0 and os.path.exists(out), "the example writes a workbook")
        cells, _ = O.read(out)
        say([s for s in cells] == SHEETS, "six sheets, in order: " + ", ".join(cells))
        bad = json.loads(json.dumps(EXAMPLE)); bad["working_capital"]["receivable_days"] = 400
        io.open(src, "w", encoding="utf-8").write(json.dumps(bad))
        out2 = os.path.join(t, "no.xlsx")
        r = subprocess.run([sys.executable, __file__, src, "-o", out2], capture_output=True, text=True)
        say(r.returncode == 1 and not os.path.exists(out2) and "B02" in r.stdout, "400 days to be paid is refused and nothing is written")
        bad = json.loads(json.dumps(EXAMPLE)); bad["illustrative"] = False
        io.open(src, "w", encoding="utf-8").write(json.dumps(bad))
        r = subprocess.run([sys.executable, __file__, src, "-o", out2], capture_output=True, text=True)
        say(r.returncode == 1 and "B03" in r.stdout, "real numbers with no sources are refused")
        io.open(src, "w", encoding="utf-8").write(json.dumps(EXAMPLE).replace('"name": ', '"name": "x", "name": ', 1))
        r = subprocess.run([sys.executable, __file__, src, "-o", out2], capture_output=True, text=True)
        say(r.returncode == 1 and "twice" in r.stdout, "a key written twice is refused")
        # the owners' money and the loan exactly cover the equipment: the start cash is nothing, not a hair below it
        even = json.loads(json.dumps(EXAMPLE))
        even["funding"].update(equity=17129.37, loan=10590.9); even["equipment"]["initial"] = 27720.27
        io.open(src, "w", encoding="utf-8").write(json.dumps(even))
        r = subprocess.run([sys.executable, __file__, src, "-o", out], capture_output=True, text=True)
        cells, _ = O.read(out)
        say(r.returncode == 0 and "below zero" not in r.stdout and cells["Balance sheet"]["B4"]["value"] == 0
            and cells["Checks"]["B11"]["value"] == "No",
            "equity and loan that exactly cover the equipment: the start cash is 0, and nothing says the model needs more money")
        # a year whose cash comes to exactly nothing is stored a hair below zero (-3.6e-12) and kept as it is (a year-end
        # cash can be fractions of a penny, and they carry into the next year); the shortfall is judged at half a penny
        yearend = json.loads(json.dumps(EXAMPLE)); yearend["years"] = 1; yearend["tax_rate"] = 0.2
        yearend["sales"].update(units=267, price=145); yearend["costs"].update(unit_cost=4, fixed_costs=15001)
        yearend["equipment"].update(initial=45000, yearly=500, depreciation_rate=0.2)
        yearend["working_capital"].update(receivable_days=0, inventory_days=0, payable_days=0)
        yearend["funding"].update(equity=27567.2, loan=31000, interest_rate=0.08, loan_years=1)
        for equity, short in ((27567.2, False), (27567.19, True)):
            yearend["funding"]["equity"] = equity
            io.open(src, "w", encoding="utf-8").write(json.dumps(yearend))
            r = subprocess.run([sys.executable, __file__, src, "-o", out], capture_output=True, text=True)
            cells, _ = O.read(out)
            said = cells["Checks"]["B11"]["value"]
            printed = next((x for x in r.stdout.splitlines() if "cash at the end of each year" in x), "")
            say(r.returncode == 0 and ("below zero in" in r.stdout) == short and said.startswith("Yes") == short
                and ("-0 (" not in printed) and (short or "year 0 (" in printed),
                f"owners' money of {equity:,.2f}, so the year ends with {'a penny short' if short else 'exactly nothing'}: "
                + ("both the note and the Checks sheet say it needs money" if short else "neither the note nor the Checks sheet says it needs money, and the cash prints as 0, not -0")
                + f" (Checks!B11 {said!r})")
        # control characters (the workbook's XML cannot hold them) and what model_check.py's K11 would refuse are refused here first
        for label, change, words in (("a star in the name", lambda b: b.update(name=b["name"] + " \u2605"), "icon"),
                                     ("a bell character in the name", lambda b: b.update(name="Harbour\u0007Bikes"), "control character"),
                                     ("fixed costs of ten trillion", lambda b: b["costs"].update(fixed_costs=1e13), "between 0 and 1e+09"),
                                     ("a 400-digit number of units", lambda b: b["sales"].update(units=10 ** 400), "digits")):
            bad = json.loads(json.dumps(EXAMPLE)); change(bad)
            io.open(src, "w", encoding="utf-8").write(json.dumps(bad))
            r = subprocess.run([sys.executable, __file__, src, "-o", out2], capture_output=True, text=True)
            say(r.returncode == 1 and not os.path.exists(out2) and words in r.stdout and "Traceback" not in r.stderr, f"{label} is refused, and nothing is written")
    print(f"make_model selftest: {sum(ok)}/{len(ok)} passed")
    return 0 if all(ok) else 2


def main(argv):
    if "--selftest" in argv:
        return selftest()
    if "--init" in argv:
        i = argv.index("--init")
        if i + 1 >= len(argv):
            print(__doc__.strip().split("\n\n")[1]); return 2
        path = argv[i + 1]
        if os.path.exists(path):
            print(f"{path} exists; not overwriting it"); return 2
        io.open(path, "w", encoding="utf-8").write(json.dumps(EXAMPLE, indent=2, ensure_ascii=False) + "\n")
        print(f"wrote {path}: an illustrative bike-repair business to edit"); return 0
    if "-o" not in argv or len(argv) != 3:
        print(__doc__.strip().split("\n\n")[1]); return 2
    i = argv.index("-o")
    if i + 1 >= len(argv):
        print(__doc__.strip().split("\n\n")[1]); return 2
    src = [a for j, a in enumerate(argv) if j not in (i, i + 1)]
    return make(src[0], argv[i + 1])


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
