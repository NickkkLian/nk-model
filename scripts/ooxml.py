#!/usr/bin/env python3
"""ooxml.py — write and read the small part of the .xlsx format a three-statement model needs, with the standard
library only: sheets, shared strings, a handful of styles, number formats, column widths, frozen panes, formulas
stored together with the value they give, and charts (clustered columns and lines) that carry their own copy of the
numbers they draw.

Why formulas are stored with their values: Excel recalculates when it opens a file, but the viewers that do not —
Quick Look, phone previews, most web previews — show only the value written next to a formula, and a formula written
without one shows as an empty cell there. make_model.py computes every value before it writes, and model_check.py
recomputes them all from the formulas and compares.

    python3 ooxml.py --selftest

Not supported, on purpose: macros, external links, merged cells, conditional formats, pivot tables, comments.
"""
import datetime as _dt
import io
import re
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def col_letter(n):
    """1 -> A, 27 -> AA"""
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def col_number(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def ref(row, col, absolute=False):
    return f"${col_letter(col)}${row}" if absolute else f"{col_letter(col)}{row}"


def split_ref(r):
    m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", r)
    if not m:
        raise ValueError(f"not a cell reference: {r!r}")
    return int(m.group(2)), col_number(m.group(1))


def quote_sheet(name):
    """a sheet name as a formula writes it: quoted when it has anything but letters, digits and _"""
    return name if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) else "'" + name.replace("'", "''") + "'"


def number_text(x):
    """a number as the file stores it: the shortest text that reads back as the same double"""
    if isinstance(x, bool):
        return "1" if x else "0"
    if isinstance(x, int) or (isinstance(x, float) and x.is_integer() and abs(x) < 1e15):
        return str(int(x))
    return repr(float(x))


# ---- styles --------------------------------------------------------------------------------------------------------
class Styles:
    """A named style is (font, fill, border, number format, horizontal alignment, indent). Each distinct one becomes one
    cellXfs entry; index 0 stays the plain default the format requires."""

    def __init__(self, body_font="Inter", size=10):
        self.fonts = [(body_font, size, False, False, "FF1D1B24")]
        self.fills = [None, "gray125"]                      # the format reserves the first two fills
        self.borders = [(None, None)]
        self.numfmts = {}                                  # code -> id (custom ids start at 164)
        self.xfs = [(0, 0, 0, 0, None, 0)]
        self.named = {}
        self.body_font, self.size = body_font, size

    def _index(self, table, item):
        if item not in table:
            table.append(item)
        return table.index(item)

    def add(self, name, font=None, fill=None, border=None, numfmt=None, align=None, indent=0):
        """font: dict(name, size, bold, italic, color) · fill: 'RRGGBB' · border: ('thin'|'medium'|None, top? bool)
        numfmt: an Excel format code · align: 'left'|'center'|'right'"""
        f = font or {}
        fi = self._index(self.fonts, (f.get("name", self.body_font), f.get("size", self.size), bool(f.get("bold")),
                                      bool(f.get("italic")), "FF" + f.get("color", "1D1B24").upper()))
        fl = 0 if not fill else self._index(self.fills, "FF" + fill.upper())
        bd = 0 if not border else self._index(self.borders, border)
        nf = 0
        if numfmt:
            builtin = {"0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10}
            nf = builtin.get(numfmt) or self.numfmts.setdefault(numfmt, 164 + len(self.numfmts))
        self.named[name] = self._index(self.xfs, (fi, fl, bd, nf, align, indent))
        return self.named[name]

    def xml(self):
        out = [f'<styleSheet xmlns="{NS}">']
        if self.numfmts:
            out.append(f'<numFmts count="{len(self.numfmts)}">' + "".join(
                f'<numFmt numFmtId="{i}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>' for code, i in self.numfmts.items()) + "</numFmts>")
        out.append(f'<fonts count="{len(self.fonts)}">')
        for name, size, bold, italic, color in self.fonts:
            out.append("<font>" + ("<b/>" if bold else "") + ("<i/>" if italic else "") +
                       f'<sz val="{size}"/><color rgb="{color}"/><name val="{escape(name)}"/><family val="2"/></font>')
        out.append("</fonts>")
        out.append(f'<fills count="{len(self.fills)}">')
        for f in self.fills:
            if f is None:
                out.append('<fill><patternFill patternType="none"/></fill>')
            elif f == "gray125":
                out.append('<fill><patternFill patternType="gray125"/></fill>')
            else:
                out.append(f'<fill><patternFill patternType="solid"><fgColor rgb="{f}"/><bgColor indexed="64"/></patternFill></fill>')
        out.append("</fills>")
        out.append(f'<borders count="{len(self.borders)}">')
        for weight, double in self.borders:
            if weight is None:
                out.append("<border><left/><right/><top/><bottom/><diagonal/></border>")
            else:
                top = f'<top style="{weight}"><color rgb="FF4F4A57"/></top>'
                bottom = '<bottom style="double"><color rgb="FF4F4A57"/></bottom>' if double else "<bottom/>"
                out.append(f"<border><left/><right/>{top}{bottom}<diagonal/></border>")
        out.append("</borders>")
        out.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        out.append(f'<cellXfs count="{len(self.xfs)}">')
        for fi, fl, bd, nf, align, indent in self.xfs:
            attrs = f'numFmtId="{nf}" fontId="{fi}" fillId="{fl}" borderId="{bd}" xfId="0"'
            attrs += (' applyNumberFormat="1"' if nf else "") + (' applyFont="1"' if fi else "") + \
                     (' applyFill="1"' if fl else "") + (' applyBorder="1"' if bd else "")
            if align or indent:
                out.append(f'<xf {attrs} applyAlignment="1"><alignment' + (f' horizontal="{align}"' if align else "") +
                           (f' indent="{indent}"' if indent else "") + "/></xf>")
            else:
                out.append(f"<xf {attrs}/>")
        out.append("</cellXfs>")
        out.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        out.append("</styleSheet>")
        return "".join(out)


# ---- the workbook ----------------------------------------------------------------------------------------------------
class Sheet:
    def __init__(self, book, name):
        if not (0 < len(name) <= 31) or re.search(r"[\[\]:*?/\\]", name):
            raise ValueError(f"not a usable sheet name: {name!r}")
        self.book, self.name = book, name
        self.cells = {}                 # (row, col) -> (kind, value, formula, style index)
        self.widths = {}
        self.freeze = None
        self.charts = []
        self.tab = None

    def set(self, row, col, value=None, formula=None, style=None):
        """value: a number, text, bool or None; formula: the formula without its '=', stored with value as its result"""
        if formula is not None and formula.startswith("="):
            formula = formula[1:]
        st = self.book.styles.named[style] if style else 0
        kind = "n" if isinstance(value, (int, float)) and not isinstance(value, bool) else \
               "b" if isinstance(value, bool) else "s" if isinstance(value, str) else "blank"
        if formula is not None and kind == "s":
            kind = "str"                # a formula whose result is text is stored inline, not in the shared strings
        self.cells[(row, col)] = (kind, value, formula, st)

    def width(self, col, chars):
        self.widths[col] = chars

    def chart(self, **spec):
        self.charts.append(spec)

    def xml(self, strings):
        rows = {}
        for (r, c), cell in self.cells.items():
            rows.setdefault(r, []).append((c, cell))
        out = [f'<worksheet xmlns="{NS}" xmlns:r="{R}">']
        if self.tab:
            out.append(f'<sheetPr><tabColor rgb="FF{self.tab}"/></sheetPr>')
        out.append("<sheetViews><sheetView workbookViewId=\"0\"" + (' tabSelected="1"' if self.book.sheets[0] is self else "") +
                   ' showGridLines="0">')
        if self.freeze:
            fr, fc = split_ref(self.freeze)
            xs, ys = fc - 1, fr - 1
            pane = "bottomRight" if xs and ys else "bottomLeft" if ys else "topRight"
            out.append(f'<pane{f" xSplit={chr(34)}{xs}{chr(34)}" if xs else ""}{f" ySplit={chr(34)}{ys}{chr(34)}" if ys else ""} '
                       f'topLeftCell="{self.freeze}" activePane="{pane}" state="frozen"/>')
        out.append("</sheetView></sheetViews>")
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        if self.widths:
            out.append("<cols>" + "".join(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>' for c, w in sorted(self.widths.items())) + "</cols>")
        out.append("<sheetData>")
        for r in sorted(rows):
            out.append(f'<row r="{r}">')
            for c, (kind, value, formula, st) in sorted(rows[r]):
                a = f'r="{col_letter(c)}{r}"' + (f' s="{st}"' if st else "")
                f = f"<f>{escape(formula)}</f>" if formula is not None else ""
                if kind == "blank":
                    out.append(f"<c {a}>{f}</c>" if f else f"<c {a}/>")
                elif kind == "n":
                    out.append(f"<c {a}>{f}<v>{number_text(value)}</v></c>")
                elif kind == "b":
                    out.append(f'<c {a} t="b">{f}<v>{1 if value else 0}</v></c>')
                elif kind == "str":
                    out.append(f'<c {a} t="str">{f}<v>{escape(value)}</v></c>')
                else:
                    out.append(f'<c {a} t="s"><v>{strings.index(value)}</v></c>')
            out.append("</row>")
        out.append("</sheetData>")
        out.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        if self.charts:
            out.append('<drawing r:id="rId1"/>')
        out.append("</worksheet>")
        return "".join(out)


class Strings:
    def __init__(self):
        self.items, self.pos = [], {}

    def index(self, s):
        if s not in self.pos:
            self.pos[s] = len(self.items)
            self.items.append(s)
        return self.pos[s]


class Workbook:
    def __init__(self, title="", creator="", body_font="Inter"):
        self.sheets, self.title, self.creator = [], title, creator
        self.styles = Styles(body_font)

    def sheet(self, name):
        s = Sheet(self, name)
        if any(x.name.lower() == name.lower() for x in self.sheets):
            raise ValueError(f"two sheets called {name!r}")
        self.sheets.append(s)
        return s

    def save(self, path, when=None):
        """write the package; `when` fixes the timestamp in docProps so two runs give the same bytes"""
        when = when or _dt.datetime(2026, 1, 1)
        strings = Strings()
        for s in self.sheets:                                  # shared strings in a stable order: sheet, row, column
            for key in sorted(s.cells):
                kind, value, formula, st = s.cells[key]
                if kind == "s":
                    strings.index(value)
        parts = {}
        chart_no = 0
        drawings = []
        for i, s in enumerate(self.sheets, 1):
            parts[f"xl/worksheets/sheet{i}.xml"] = s.xml(strings)
            if s.charts:
                d = len(drawings) + 1
                drawings.append(d)
                parts[f"xl/worksheets/_rels/sheet{i}.xml.rels"] = (
                    f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" '
                    f'Type="{R}/drawing" Target="../drawings/drawing{d}.xml"/></Relationships>')
                anchors, rels = [], []
                for k, spec in enumerate(s.charts, 1):
                    chart_no += 1
                    parts[f"xl/charts/chart{chart_no}.xml"] = chart_xml(spec)
                    rels.append(f'<Relationship Id="rId{k}" Type="{R}/chart" Target="../charts/chart{chart_no}.xml"/>')
                    anchors.append(anchor_xml(spec["at"], k, spec.get("title", "chart")))
                parts[f"xl/drawings/drawing{d}.xml"] = (f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}">' + "".join(anchors) + "</xdr:wsDr>")
                parts[f"xl/drawings/_rels/drawing{d}.xml.rels"] = f'<Relationships xmlns="{PKG}">' + "".join(rels) + "</Relationships>"
        n = len(self.sheets)
        parts["xl/workbook.xml"] = (
            f'<workbook xmlns="{NS}" xmlns:r="{R}"><workbookPr date1904="0"/>'
            f'<bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
            "".join(f'<sheet name="{escape(s.name, {chr(34): "&quot;"})}" sheetId="{i}" r:id="rId{i}"/>' for i, s in enumerate(self.sheets, 1)) +
            '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        parts["xl/_rels/workbook.xml.rels"] = (
            f'<Relationships xmlns="{PKG}">' +
            "".join(f'<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, n + 1)) +
            f'<Relationship Id="rId{n + 1}" Type="{R}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId{n + 2}" Type="{R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
        parts["xl/styles.xml"] = self.styles.xml()
        parts["xl/sharedStrings.xml"] = (f'<sst xmlns="{NS}" count="{len(strings.items)}" uniqueCount="{len(strings.items)}">' +
                                         "".join(f'<si><t xml:space="preserve">{escape(t)}</t></si>' for t in strings.items) + "</sst>")
        stamp = when.strftime("%Y-%m-%dT%H:%M:%SZ")
        parts["docProps/core.xml"] = (
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{escape(self.title)}</dc:title><dc:creator>{escape(self.creator)}</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{stamp}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{stamp}</dcterms:modified></cp:coreProperties>')
        parts["docProps/app.xml"] = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                                     "<Application>nk-model</Application></Properties>")
        parts["_rels/.rels"] = (
            f'<Relationships xmlns="{PKG}">'
            f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            f'<Relationship Id="rId3" Type="{R}/extended-properties" Target="docProps/app.xml"/></Relationships>')
        ov = [('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
              ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
              ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'),
              ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
              ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')]
        ov += [(f'/xl/worksheets/sheet{i}.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml') for i in range(1, n + 1)]
        ov += [(f'/xl/drawings/drawing{d}.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml') for d in drawings]
        ov += [(f'/xl/charts/chart{k}.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml') for k in range(1, chart_no + 1)]
        parts["[Content_Types].xml"] = (
            f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>' +
            "".join(f'<Override PartName="{p}" ContentType="{t}"/>' for p, t in ov) + "</Types>")
        order = ["[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml", "xl/workbook.xml",
                 "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml"]
        order += sorted(p for p in parts if p not in order)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for p in order:
                info = zipfile.ZipInfo(p, date_time=(when.year, when.month, when.day, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + parts[p])
        with open(path, "wb") as f:
            f.write(buf.getvalue())
        return len(buf.getvalue())


# ---- charts --------------------------------------------------------------------------------------------------------
def anchor_xml(at, rid, name):
    """at = (first col, first row, last col, last row), 1-based and inclusive of the first cell"""
    c0, r0, c1, r1 = at
    return (f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{c0 - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r0 - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{c1 - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1 - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{rid + 1}" name="{escape(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>'
            f'<a:graphicData uri="{C}"><c:chart xmlns:c="{C}" xmlns:r="{R}" r:id="rId{rid}"/></a:graphicData>'
            '</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')


def _rich(text, size=1200, bold=True, color="1D1B24"):
    return (f'<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{size}" b="{1 if bold else 0}"/></a:pPr>'
            f'<a:r><a:rPr lang="en-GB" sz="{size}" b="{1 if bold else 0}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr>'
            f'<a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx>')


def chart_xml(spec):
    """spec: kind ('bar'|'line'), title, categories (ref, [labels]), series [(name, ref, [values], 'RRGGBB')],
    number_format, at. Every reference carries a cache of the values, so a viewer that does not recalculate draws it."""
    kind, series = spec["kind"], spec["series"]
    cat_ref, cats = spec["categories"]
    fmt = escape(spec.get("number_format", "#,##0"), {chr(34): "&quot;"})

    def cat_xml():
        return (f'<c:cat><c:strRef><c:f>{escape(cat_ref)}</c:f><c:strCache><c:ptCount val="{len(cats)}"/>' +
                "".join(f'<c:pt idx="{i}"><c:v>{escape(str(v))}</c:v></c:pt>' for i, v in enumerate(cats)) +
                "</c:strCache></c:strRef></c:cat>")

    body = []
    for i, (name, vref, vals, color) in enumerate(series):
        shape = (f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' if kind == "bar" else
                 f'<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:round/></a:ln></c:spPr>'
                 f'<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                 f'<a:ln><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr></c:marker>')
        body.append(f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:v>{escape(name)}</c:v></c:tx>{shape}' + cat_xml() +
                    f'<c:val><c:numRef><c:f>{escape(vref)}</c:f><c:numCache><c:formatCode>{fmt}</c:formatCode><c:ptCount val="{len(vals)}"/>' +
                    "".join(f'<c:pt idx="{j}"><c:v>{number_text(v)}</c:v></c:pt>' for j, v in enumerate(vals)) +
                    "</c:numCache></c:numRef></c:val>" + ('<c:smooth val="0"/>' if kind == "line" else "") + "</c:ser>")
    plot = (f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{"".join(body)}'
            '<c:gapWidth val="80"/><c:overlap val="-10"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>' if kind == "bar" else
            f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{"".join(body)}<c:marker val="1"/>'
            '<c:axId val="1001"/><c:axId val="1002"/></c:lineChart>')
    grid = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E4DCDC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
    axis_text = ('<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="4F4A57"/></a:solidFill>'
                 '</a:defRPr></a:pPr><a:endParaRPr lang="en-GB"/></a:p></c:txPr>')
    line = '<c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="BFB6B6"/></a:solidFill></a:ln></c:spPr>'
    return (f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}"><c:roundedCorners val="0"/><c:chart>'
            f'<c:title>{_rich(spec.get("title", ""))}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>{plot}'
            f'<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>'
            f'<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>{line}{axis_text}'
            '<c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
            f'<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>{grid}'
            f'<c:numFmt formatCode="{fmt}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
            f'<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>{axis_text}<c:crossAx val="1001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
            '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + axis_text + '</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>')


# ---- reading ---------------------------------------------------------------------------------------------------------
BUILTIN_FORMATS = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
                   12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
                   19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)",
                   38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
                   46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@"}


def shift_formula(text, dr, dc):
    """a shared formula as it reads in a cell dr rows and dc columns from the cell that holds it: every reference part
    without a $ moves by that much (Excel stores a filled-down run of formulas once, as a shared formula)"""
    import xlformula as X
    out, pos = [], 0

    def move(part):
        m = re.fullmatch(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", part)
        c = m.group(2) if m.group(1) else col_letter(col_number(m.group(2)) + dc)
        r = m.group(4) if m.group(3) else str(int(m.group(4)) + dr)
        return m.group(1) + c + m.group(3) + r
    while pos < len(text):
        m = X.TOKEN.match(text, pos)
        if not m or m.end() == pos:
            out.append(text[pos:]); break
        if m.lastgroup == "ref":
            start, end = m.span("ref")
            ref_text = m.group("ref")
            sheet, bang, cells = ref_text.rpartition("!")
            out.append(text[pos:start] + sheet + bang + ":".join(move(x) for x in cells.split(":")))
            pos = end
        else:
            out.append(text[pos:m.end()]); pos = m.end()
    return "".join(out)


def read(path):
    """{sheet name: {ref: {"value": ..., "formula": str|None, "style": int, "numfmt": the cell's number format code}}},
    plus the workbook's parts, from any .xlsx. A shared formula is given to each cell it covers, moved to that cell."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError("a part appears twice in the package")
        parts = {n: z.read(n) for n in names}
    ns = {"m": NS}
    strings = []
    if "xl/sharedStrings.xml" in parts:
        for si in ET.fromstring(parts["xl/sharedStrings.xml"]).findall("m:si", ns):
            strings.append("".join(t.text or "" for t in si.iter("{%s}t" % NS)))
    wb = ET.fromstring(parts["xl/workbook.xml"])
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(parts["xl/_rels/workbook.xml.rels"])}
    codes = []                                           # cellXfs index -> number format code
    if "xl/styles.xml" in parts:
        st = ET.fromstring(parts["xl/styles.xml"])
        custom = {int(x.get("numFmtId")): x.get("formatCode") for x in st.iter("{%s}numFmt" % NS)}
        xfs = st.find("m:cellXfs", ns)
        for xf in (xfs if xfs is not None else []):
            i = int(xf.get("numFmtId", "0"))
            codes.append(custom.get(i, BUILTIN_FORMATS.get(i, f"(format {i})")))
    out = {}
    for s in wb.find("m:sheets", ns):
        target = rels[s.get("{%s}id" % R)].lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        cells, shared = {}, {}
        sheet_xml = ET.fromstring(parts[target])
        for c in sheet_xml.iter("{%s}c" % NS):          # the shared formulas' own cells first: a formula is stored once
            f = c.find("m:f", ns)
            if f is not None and f.get("t") == "shared" and f.text:
                shared[f.get("si")] = (f.text, *split_ref(c.get("r")))
        for c in sheet_xml.iter("{%s}c" % NS):
            t, v, f = c.get("t"), c.find("m:v", ns), c.find("m:f", ns)
            text = v.text if v is not None else None
            if t == "s" and text is not None:
                value = strings[int(text)]
            elif t in ("str", "inlineStr"):
                value = text if t == "str" else "".join(x.text or "" for x in c.iter("{%s}t" % NS))
            elif t == "b":
                value = text == "1"
            elif t == "e":
                value = "#ERROR:" + (text or "")
            elif text is not None:
                value = float(text)
            else:
                value = None
            formula = f.text if f is not None else None
            if f is not None and f.get("t") == "shared" and not f.text and f.get("si") in shared:
                text, r0, c0 = shared[f.get("si")]
                r1, c1 = split_ref(c.get("r"))
                formula = shift_formula(text, r1 - r0, c1 - c0)
            style = int(c.get("s", "0"))
            cells[c.get("r")] = {"value": value, "formula": formula, "style": style,
                                 "numfmt": codes[style] if style < len(codes) else "General"}
        out[s.get("name")] = cells
    return out, parts


def selftest():
    import os, tempfile
    ok = []

    def say(cond, text):
        ok.append(bool(cond)); print(f"  {'PASS' if cond else 'FAIL'}  {text}")

    say([col_letter(n) for n in (1, 26, 27, 52, 703)] == ["A", "Z", "AA", "AZ", "AAA"], "column letters: 1 A, 26 Z, 27 AA, 52 AZ, 703 AAA")
    say(all(col_number(col_letter(n)) == n for n in range(1, 2000)), "letters and numbers round-trip for the first 1999 columns")
    say(quote_sheet("Cash flow") == "'Cash flow'" and quote_sheet("Checks") == "Checks" and quote_sheet("Bob's") == "'Bob''s'", "sheet names are quoted in formulas when they need it")
    say([number_text(x) for x in (3, 3.0, 0.1, -2.5, True)] == ["3", "3", "0.1", "-2.5", "1"], "numbers are stored in their shortest exact form")
    wb = Workbook(title="t", creator="selftest")
    wb.styles.add("num", numfmt="#,##0.00;(#,##0.00);\"–\"", align="right")
    wb.styles.add("head", font={"bold": True, "color": "F7F0E9"}, fill="053333")
    s = wb.sheet("Data & notes")
    s.set(1, 1, "Year", style="head"); s.set(1, 2, "Value", style="head")
    for i in range(3):
        s.set(2 + i, 1, f"Y{i + 1}"); s.set(2 + i, 2, 10.5 * (i + 1), style="num")
    s.set(5, 2, 63.0, formula="SUM(B2:B4)", style="num")
    s.set(6, 2, "yes", formula='IF(B5>0,"yes","no")')
    s.set(7, 2, True, formula="B5>0")
    s.freeze = "B2"
    s.chart(kind="bar", title="Value by year", at=(4, 2, 10, 16), categories=("'Data & notes'!$A$2:$A$4", ["Y1", "Y2", "Y3"]),
            series=[("Value", "'Data & notes'!$B$2:$B$4", [10.5, 21.0, 31.5], "CA6980")])
    s.chart(kind="line", title="Again", at=(4, 17, 10, 30), categories=("'Data & notes'!$A$2:$A$4", ["Y1", "Y2", "Y3"]),
            series=[("Value", "'Data & notes'!$B$2:$B$4", [10.5, 21.0, 31.5], "3CA77A")])
    with tempfile.TemporaryDirectory() as t:
        p1, p2 = os.path.join(t, "a.xlsx"), os.path.join(t, "b.xlsx")
        wb.save(p1); wb.save(p2)
        say(open(p1, "rb").read() == open(p2, "rb").read(), "the same workbook saved twice gives the same bytes")
        cells, parts = read(p1)
        d = cells["Data & notes"]
        say(d["B5"]["formula"] == "SUM(B2:B4)" and d["B5"]["value"] == 63.0, "a formula is stored with its value, and both read back")
        say(d["B6"]["value"] == "yes" and d["B6"]["formula"].startswith("IF(") and d["B7"]["value"] is True, "text and true/false results of formulas read back")
        say(d["A2"]["value"] == "Y1" and d["B3"]["value"] == 21.0 and d["B3"]["formula"] is None, "plain text and numbers read back, with no formula")
        for name, blob in parts.items():
            if name.endswith(".xml") or name.endswith(".rels"):
                ET.fromstring(blob)                             # raises if any part is not well-formed
        say(True, f"all {len(parts)} parts are well-formed XML")
        say("xl/charts/chart1.xml" in parts and "xl/charts/chart2.xml" in parts and b"<c:numCache>" in parts["xl/charts/chart1.xml"],
            "the charts are in the package and carry a cache of the numbers they draw")
        ct = parts["[Content_Types].xml"].decode()
        say(all(f"/{p}" in ct for p in parts if p.startswith("xl/") and p.endswith(".xml") and "_rels" not in p),
            "every xl/ part has a content type")
        say(b'fullCalcOnLoad="1"' in parts["xl/workbook.xml"], "Excel is told to recalculate everything when it opens the file")
    try:
        Workbook().sheet("bad/name"); refused = False
    except ValueError:
        refused = True
    say(refused, "a sheet name Excel would refuse is refused here")
    print(f"ooxml selftest: {sum(ok)}/{len(ok)} passed")
    return 0 if all(ok) else 2


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    print(__doc__.strip().split("\n\n")[0]); sys.exit(2)
