#!/usr/bin/env python3
"""preview.py — look at a model workbook without a spreadsheet program: one HTML file, every sheet as a table of the
values stored in the file (in the sheet's own number formats, fills and weights), and each chart drawn from the copy of
the numbers the chart carries.

    python3 preview.py model.xlsx -o preview.html      # open it; add #s3 to the address to see the third sheet
    python3 preview.py --selftest

It shows what the file holds and recalculates nothing: a stored value that disagrees with its formula shows here as
stored — model_check.py is what finds that. Nothing is fetched: fonts fall back to the system's.
"""
import html, os, re, sys
from decimal import Decimal, ROUND_HALF_UP
from xml.etree import ElementTree as ET

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml as O

BUILTIN = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%"}


def sections(code):
    """split a number format on ; outside quotes"""
    out, cur, q = [], "", False
    for ch in code:
        if ch == '"':
            q = not q
        if ch == ";" and not q:
            out.append(cur); cur = ""
        else:
            cur += ch
    return out + [cur]


def fmt_number(x, code):
    """a number as the format shows it — the formats this skill writes: thousands separators, decimals, %, brackets
    for negatives and a literal for zero"""
    if code in ("General", None):
        return f"{x:,.10g}" if abs(x) >= 1e-4 or x == 0 else f"{x:.4g}"
    parts = sections(code)
    if x < 0 and len(parts) >= 2:
        part, x = parts[1], -x
    elif x == 0 and len(parts) >= 3:
        part = parts[2]
    else:
        part = parts[0]
    lit = re.fullmatch(r'"([^"]*)"', part)
    if lit:
        return lit.group(1)
    pct = "%" in part
    if pct:
        x *= 100
    m = re.search(r"[#0][#0,]*(?:\.(0+))?", part)
    decimals = len(m.group(1)) if m and m.group(1) else 0
    # halves go away from zero, as a spreadsheet shows them (Python's format() would round 1234.5 to 1234)
    x = float(Decimal(repr(x)).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP))
    body = f"{x:,.{decimals}f}" if "," in (m.group(0) if m else "") else f"{x:.{decimals}f}"
    return (part[:m.start()] if m else "") .replace('"', "") + body + (part[m.end():] if m else "").replace('"', "")


def styles(parts):
    root = ET.fromstring(parts["xl/styles.xml"])
    ns = {"m": O.NS}
    codes = dict(BUILTIN)
    for nf in root.findall("m:numFmts/m:numFmt", ns):
        codes[int(nf.get("numFmtId"))] = nf.get("formatCode")
    fonts = []
    for f in root.findall("m:fonts/m:font", ns):
        color = f.find("m:color", ns)
        fonts.append({"bold": f.find("m:b", ns) is not None, "italic": f.find("m:i", ns) is not None,
                      "color": ("#" + color.get("rgb")[-6:]) if color is not None and color.get("rgb") else None,
                      "size": float(f.find("m:sz", ns).get("val")) if f.find("m:sz", ns) is not None else 10,
                      "name": f.find("m:name", ns).get("val") if f.find("m:name", ns) is not None else ""})
    fills = []
    for fl in root.findall("m:fills/m:fill", ns):
        fg = fl.find("m:patternFill/m:fgColor", ns)
        fills.append(("#" + fg.get("rgb")[-6:]) if fg is not None and fg.get("rgb") else None)
    borders = []
    for b in root.findall("m:borders/m:border", ns):
        top, bottom = b.find("m:top", ns), b.find("m:bottom", ns)
        borders.append((top is not None and top.get("style"), bottom is not None and bottom.get("style")))
    xfs = []
    for x in root.findall("m:cellXfs/m:xf", ns):
        al = x.find("m:alignment", ns)
        xfs.append({"font": fonts[int(x.get("fontId", 0))], "fill": fills[int(x.get("fillId", 0))],
                    "border": borders[int(x.get("borderId", 0))], "code": codes.get(int(x.get("numFmtId", 0)), "General"),
                    "align": al.get("horizontal") if al is not None else None, "indent": int(al.get("indent", 0)) if al is not None else 0})
    return xfs


def charts_of(parts):
    """[(sheet index, [chart dict])] from the drawings: kind, title, categories, series (name, values, colour)"""
    ns = {"c": O.C, "a": O.A}
    out = {}
    wb_rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(parts["xl/_rels/workbook.xml.rels"])}
    sheets = ET.fromstring(parts["xl/workbook.xml"]).find("{%s}sheets" % O.NS)
    for k, s in enumerate(sheets):
        target = "xl/" + wb_rels[s.get("{%s}id" % O.R)].lstrip("/").replace("xl/", "")
        rels = target.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rels not in parts:
            continue
        for rel in ET.fromstring(parts[rels]):
            if not rel.get("Type", "").endswith("/drawing"):
                continue
            drawing = "xl/drawings/" + os.path.basename(rel.get("Target"))
            drels = drawing.replace("drawings/", "drawings/_rels/") + ".rels"
            for crel in ET.fromstring(parts[drels]):
                root = ET.fromstring(parts["xl/charts/" + os.path.basename(crel.get("Target"))])
                kind = "bar" if root.find(".//c:barChart", ns) is not None else "line"
                title = "".join(t.text or "" for t in root.findall("c:chart/c:title//a:t", ns))
                series = []
                for ser in root.iter("{%s}ser" % O.C):
                    name = ser.find("c:tx/c:v", ns)
                    cats = [p.find("c:v", ns).text for p in ser.findall("c:cat/c:strRef/c:strCache/c:pt", ns)]
                    vals = [float(p.find("c:v", ns).text) for p in ser.findall("c:val/c:numRef/c:numCache/c:pt", ns)]
                    col = ser.find(".//a:srgbClr", ns)
                    series.append({"name": name.text if name is not None else "", "cats": cats, "vals": vals,
                                   "color": "#" + (col.get("val") if col is not None else "666666")})
                out.setdefault(k, []).append({"kind": kind, "title": title, "series": series})
    return out


def nice(lo, hi, want=4):
    import math
    if hi <= lo:
        hi = lo + 1
    raw = (hi - lo) / want
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 5, 10) if raw <= m * mag)
    return [i * step for i in range(math.floor(lo / step), math.ceil(hi / step) + 1)]


def svg_chart(ch, w=640, h=300):
    cats = ch["series"][0]["cats"] if ch["series"] else []
    allv = [v for s in ch["series"] for v in s["vals"]] + [0]
    ticks = nice(min(allv), max(allv))
    y0, y1 = ticks[0], ticks[-1]
    L, R, T, B = 64, 12, 16, 48
    sx = lambda i: L + (w - L - R) * (i + 0.5) / max(1, len(cats))
    sy = lambda y: T + (h - T - B) * (1 - (y - y0) / (y1 - y0))
    out = [f'<svg viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" aria-label="{html.escape(ch["title"])}">']
    for t in ticks:
        out.append(f'<line x1="{L}" x2="{w - R}" y1="{sy(t):.1f}" y2="{sy(t):.1f}" class="{"axis" if t == 0 else "grid"}"/>'
                   f'<text x="{L - 8}" y="{sy(t) + 4:.1f}" text-anchor="end">{t:,.0f}</text>')
    for i, c in enumerate(cats):
        out.append(f'<text x="{sx(i):.1f}" y="{h - B + 18}" text-anchor="middle">{html.escape(c)}</text>')
    n = len(ch["series"])
    band = (w - L - R) / max(1, len(cats))
    for k, s in enumerate(ch["series"]):
        if ch["kind"] == "bar":
            bw = band * 0.7 / max(1, n)
            for i, v in enumerate(s["vals"]):
                x = sx(i) - band * 0.35 + k * bw
                top, bot = sorted((sy(v), sy(0)))
                out.append(f'<rect x="{x:.1f}" y="{top:.1f}" width="{bw - 2:.1f}" height="{max(0.5, bot - top):.1f}" fill="{s["color"]}"/>')
        else:
            pts = " ".join(f"{sx(i):.1f},{sy(v):.1f}" for i, v in enumerate(s["vals"]))
            out.append(f'<polyline points="{pts}" fill="none" stroke="{s["color"]}" stroke-width="2.5"/>' +
                       "".join(f'<circle cx="{sx(i):.1f}" cy="{sy(v):.1f}" r="4" fill="{s["color"]}"/>' for i, v in enumerate(s["vals"])))
    lx = L
    for s in ch["series"]:
        out.append(f'<rect x="{lx}" y="{h - 16}" width="10" height="10" fill="{s["color"]}"/><text x="{lx + 14}" y="{h - 7}">{html.escape(s["name"])}</text>')
        lx += 24 + 8 * len(s["name"])
    return "".join(out) + "</svg>"


def render(path):
    cells, parts = O.read(path)
    xfs = styles(parts)
    charts = charts_of(parts)
    names = list(cells)
    widths = {}
    for k in range(1, len(names) + 1):
        root = ET.fromstring(parts[f"xl/worksheets/sheet{k}.xml"])
        widths[k - 1] = {int(c.get("min")): float(c.get("width")) for c in root.iter("{%s}col" % O.NS)}
    css = """
  :root { --band: #053333; --ink: #1d1b24; --muted: #6e6674; --line: #e4dcdc; --bg: #fbf6f4; }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--ink); font: 14px/1.45 Inter, "Helvetica Neue", Arial, system-ui, sans-serif; }
  nav { display: flex; gap: 2px; padding: 10px 16px 0; background: var(--band); overflow-x: auto; }
  nav a { color: #f7f0e9; text-decoration: none; padding: 8px 14px; border-radius: 6px 6px 0 0; font-size: 13px; white-space: nowrap; }
  nav a:hover { background: rgba(247,240,233,.14); }
  section { display: none; padding: 20px 16px 40px; }
  section:target, body:not(:has(section:target)) section:first-of-type { display: block; }
  table { border-collapse: collapse; }
  td { padding: 3px 10px; white-space: nowrap; vertical-align: bottom; }
  td.num { text-align: right; font-family: "Space Mono", ui-monospace, Menlo, monospace; font-size: 13px; }
  svg { max-width: 100%; height: auto; display: block; margin: 8px 0 24px; }
  svg text { font: 12px ui-monospace, Menlo, monospace; fill: #4f4a57; }
  svg .grid { stroke: var(--line); } svg .axis { stroke: #8a8290; }
  .stamp { color: var(--muted); font-size: 12px; padding: 0 16px 24px; }
"""
    out = ["<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">",
           f"<title>{html.escape(names[0] if names else 'workbook')}: preview</title><style>{css}</style></head><body>",
           "<nav>" + "".join(f'<a href="#s{k + 1}">{html.escape(n)}</a>' for k, n in enumerate(names)) + "</nav>"]
    for k, name in enumerate(names):
        out.append(f'<section id="s{k + 1}" aria-label="{html.escape(name)}">')
        cs = cells[name]
        pos = {O.split_ref(r): c for r, c in cs.items()}
        if pos:
            rmax, cmax = max(r for r, _ in pos), max(c for _, c in pos)
            out.append("<table>")
            for r in range(1, rmax + 1):
                row = []
                skip = 0
                for c in range(1, cmax + 1):
                    if skip:
                        skip -= 1; continue
                    cell = pos.get((r, c))
                    # text runs on over the empty cells to its right, as a spreadsheet shows it, instead of widening its column
                    span = 1
                    if cell and isinstance(cell["value"], str) and c > 1:
                        while c + span <= cmax and not pos.get((r, c + span)):
                            span += 1
                        skip = span - 1
                    st = xfs[cell["style"]] if cell else xfs[0]
                    v = cell["value"] if cell else None
                    css_cell = []
                    w = widths[k].get(c)
                    if w and r == 1:
                        css_cell.append(f"min-width:{w * 7:.0f}px")
                    f = st["font"]
                    if f["bold"]:
                        css_cell.append("font-weight:700")
                    if f["italic"]:
                        css_cell.append("font-style:italic")
                    if f["color"] and f["color"].lower() != "#1d1b24":
                        css_cell.append(f"color:{f['color']}")
                    if f["size"] and f["size"] != 10:
                        css_cell.append(f"font-size:{f['size'] * 1.4:.0f}px")
                    if "Fraunces" in f["name"]:
                        css_cell.append('font-family:Fraunces,Georgia,"Times New Roman",serif')
                    if st["fill"]:
                        css_cell.append(f"background:{st['fill']}")
                    top, bottom = st["border"]
                    if top:
                        css_cell.append("border-top:1px solid #4f4a57")
                    if bottom:
                        css_cell.append("border-bottom:3px double #4f4a57")
                    if st["indent"]:
                        css_cell.append(f"padding-left:{10 + 12 * st['indent']}px")
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        text, cls = fmt_number(float(v), st["code"]), "num"
                    else:
                        text, cls = ("" if v is None else str(v)), ("num" if st["align"] == "right" else "")
                    cs = f' colspan="{span}"' if span > 1 else ""
                    row.append(f'<td{cs} class="{cls}" style="{";".join(css_cell)}">{html.escape(text)}</td>' if (css_cell or cls) else f"<td{cs}>{html.escape(text)}</td>")
                out.append("<tr>" + "".join(row) + "</tr>")
            out.append("</table>")
        for ch in charts.get(k, []):
            out.append(f"<h3>{html.escape(ch['title'])}</h3>" + svg_chart(ch))
        out.append("</section>")
    out.append(f'<p class="stamp">Preview of {html.escape(os.path.basename(path))}: the values stored in the file, not recalculated. '
               "Charts are drawn from the numbers each chart carries.</p></body></html>")
    return "\n".join(out)


def selftest():
    import tempfile, subprocess
    here = os.path.dirname(os.path.abspath(__file__))
    ok = []

    def say(cond, text):
        ok.append(bool(cond)); print(f"  {'PASS' if cond else 'FAIL'}  {text}")

    say(fmt_number(-1234.5, '#,##0;(#,##0);"–"') == "(1,235)" and fmt_number(0, '#,##0;(#,##0);"–"') == "–", "brackets for a negative, the literal for zero")
    say(fmt_number(0.12, "0.0%") == "12.0%" and fmt_number(2400, "#,##0") == "2,400" and fmt_number(45, "#,##0.00") == "45.00", "percent, thousands and decimals")
    say(fmt_number(0.0, "0.00;-0.00;0.00") == "0.00" and fmt_number(-3.5, "0.00;-0.00;0.00") == "-3.50", "the check format: 0.00, and a real difference keeps its sign")
    with tempfile.TemporaryDirectory() as t:
        src, xl = os.path.join(t, "b.json"), os.path.join(t, "m.xlsx")
        subprocess.run([sys.executable, os.path.join(here, "make_model.py"), "--init", src], capture_output=True)
        subprocess.run([sys.executable, os.path.join(here, "make_model.py"), src, "-o", xl], capture_output=True)
        page = render(xl)
        cells, parts = O.read(xl)
        say(page.count("<section ") == len(cells) == 6, "one section per sheet")
        net = cells["Income statement"]["B16"]["value"]
        say(f"{net:,.0f}" in page, f"a stored value appears as the sheet formats it (net profit {net:,.0f})")
        say(page.count("<svg ") == 2 and "<rect" in page and "<polyline" in page, "both charts are drawn, bars and a line")
        say("http://" not in page.replace("http://www.w3.org", "") and "https://" not in page, "the page fetches nothing")
        drawn = charts_of(parts)[5][0]["series"][0]["vals"]
        stored = [cells["Income statement"][f"{c}6"]["value"] for c in "BCD"]
        say(drawn == stored, f"the revenue bars are drawn from the chart's own copy, which holds the stored revenue ({len(drawn)} years)")
        first = page.split("<svg ")[1].split("</svg>")[0]
        bars = first.count('<rect x="') - 2                 # the bar chart's rects, less its two legend keys
        say(bars == 6, f"six bars: two series over three years ({bars})")
    print(f"preview selftest: {sum(ok)}/{len(ok)} passed")
    return 0 if all(ok) else 2


def main(argv):
    if "--selftest" in argv:
        return selftest()
    if "-o" not in argv or len(argv) != 3:
        print(__doc__.strip().split("\n\n")[1]); return 2
    i = argv.index("-o")
    src = [a for j, a in enumerate(argv) if j not in (i, i + 1)][0]
    page = render(src)
    open(argv[i + 1], "w", encoding="utf-8").write(page)
    print(f"wrote {argv[i + 1]}: {page.count('<section ')} sheets; open it, or add #s2, #s3 ... to the address for the others")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
