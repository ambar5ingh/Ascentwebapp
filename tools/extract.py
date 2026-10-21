#!/usr/bin/env python3
"""
Extract everything the ASCENT web app needs from the ASCENT .xlsm workbook.

Reads the Open XML parts directly (cells, formulas, styles, data validation,
conditional formatting, drawings/shapes/buttons, images, charts, comments,
tables, defined names, form controls, external-link caches) and writes compact
JSON files that the browser app loads.

Usage:  python tools/extract.py excel/WRI_India_ASCENT_Beta_V25.xlsm site/data
"""
import colorsys
import json
import os
import posixpath
import re
import shutil
import sys
import zipfile

from lxml import etree

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "tc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
}
M = "{%s}" % NS["m"]
R_ID = "{%s}id" % NS["r"]
R_EMBED = "{%s}embed" % NS["r"]

INDEXED = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    "000000", "FFFFFF",
]

ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!"}


# --------------------------------------------------------------------------- helpers
def col2num(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def num2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


CELL_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def parse_cell(ref):
    m = CELL_RE.match(ref)
    return int(m.group(2)), col2num(m.group(1))


def parse_range(ref):
    ref = ref.replace("$", "")
    if ":" in ref:
        a, b = ref.split(":")
        if a.isdigit():  # whole rows
            return [int(a), 1, int(b), 16384]
        if a.isalpha():  # whole columns
            return [1, col2num(a), 1048576, col2num(b)]
        r1, c1 = parse_cell(a)
        r2, c2 = parse_cell(b)
        return [min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)]
    r, c = parse_cell(ref)
    return [r, c, r, c]


def sqref_list(s):
    return [parse_range(x) for x in (s or "").split() if x]


class Pkg:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())

    def has(self, p):
        return p in self.names

    def xml(self, p):
        return etree.fromstring(self.z.read(p))

    def raw(self, p):
        return self.z.read(p)

    def rels(self, part):
        d, f = posixpath.split(part)
        rp = posixpath.join(d, "_rels", f + ".rels")
        out = {}
        if not self.has(rp):
            return out
        for rel in self.xml(rp).findall("{%s}Relationship" % NS["rel"]):
            tgt = rel.get("Target")
            if rel.get("TargetMode") == "External":
                out[rel.get("Id")] = {"type": rel.get("Type").split("/")[-1], "target": tgt, "external": True}
            elif tgt.startswith("#"):
                out[rel.get("Id")] = {"type": rel.get("Type").split("/")[-1], "target": tgt, "external": False}
            else:
                full = posixpath.normpath(posixpath.join(d, tgt)) if not tgt.startswith("/") else tgt.lstrip("/")
                out[rel.get("Id")] = {"type": rel.get("Type").split("/")[-1], "target": full, "external": False}
        return out


# --------------------------------------------------------------------------- colours
class Colors:
    def __init__(self, pkg):
        self.theme = []
        self.scheme = {}
        if pkg.has("xl/theme/theme1.xml"):
            t = pkg.xml("xl/theme/theme1.xml")
            cs = t.find(".//a:clrScheme", NS)
            for el in cs:
                tag = etree.QName(el).localname
                c = el[0]
                val = c.get("lastClr") if etree.QName(c).localname == "sysClr" else c.get("val")
                self.scheme[tag] = val
            order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
            self.theme = [self.scheme.get(k, "000000") for k in order]
            self.scheme.update({"bg1": self.scheme.get("lt1"), "tx1": self.scheme.get("dk1"),
                                "bg2": self.scheme.get("lt2"), "tx2": self.scheme.get("dk2")})
        self.indexed = list(INDEXED)

    @staticmethod
    def tint(hexc, t):
        r, g, b = (int(hexc[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if t < 0:
            l = l * (1 + t)
        else:
            l = l * (1 - t) + t
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))

    def sml(self, el):
        """SpreadsheetML colour element (fgColor, color, ...) -> 'RRGGBB' or None."""
        if el is None:
            return None
        if el.get("auto") == "1":
            return None
        c = None
        if el.get("rgb"):
            c = el.get("rgb")[-6:]
        elif el.get("theme") is not None:
            i = int(el.get("theme"))
            c = self.theme[i] if i < len(self.theme) else "000000"
        elif el.get("indexed") is not None:
            i = int(el.get("indexed"))
            if i == 64:
                return None
            c = self.indexed[i] if i < len(self.indexed) else None
        if c and el.get("tint"):
            c = self.tint(c, float(el.get("tint")))
        return c

    def dml(self, parent):
        """DrawingML colour holder (solidFill etc) -> ('RRGGBB', alpha) or None."""
        if parent is None:
            return None
        for el in parent:
            tag = etree.QName(el).localname
            base = None
            if tag == "srgbClr":
                base = el.get("val")
            elif tag == "schemeClr":
                base = self.scheme.get(el.get("val"), None)
                if el.get("val") == "phClr":
                    base = None
            elif tag == "sysClr":
                base = el.get("lastClr") or "000000"
            elif tag == "prstClr":
                base = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "blue": "0000FF", "green": "008000",
                        "yellow": "FFFF00", "gray": "808080"}.get(el.get("val"), "000000")
            elif tag == "scrgbClr":
                base = "%02X%02X%02X" % tuple(round(int(el.get(k)) / 100000 * 255) for k in ("r", "g", "b"))
            else:
                continue
            if base is None:
                return {"ph": True, "mods": self._mods(el)}
            return self.apply_mods(base, self._mods(el))
        return None

    @staticmethod
    def _mods(el):
        return [(etree.QName(m).localname, int(m.get("val", "0"))) for m in el]

    def apply_mods(self, base, mods):
        r, g, b = (int(base[i:i + 2], 16) / 255 for i in (0, 2, 4))
        alpha = 1.0
        for name, v in mods:
            v = v / 100000
            if name in ("lumMod", "lumOff"):
                h, l, s = colorsys.rgb_to_hls(r, g, b)
                l = l * v if name == "lumMod" else min(1, l + v)
                r, g, b = colorsys.hls_to_rgb(h, l, s)
            elif name == "tint":
                r, g, b = (x + (1 - x) * (1 - v) for x in (r, g, b))
            elif name == "shade":
                r, g, b = (x * v for x in (r, g, b))
            elif name == "alpha":
                alpha = v
            elif name == "satMod":
                h, l, s = colorsys.rgb_to_hls(r, g, b)
                r, g, b = colorsys.hls_to_rgb(h, l, min(1, s * v))
        r, g, b = (max(0, min(1, x)) for x in (r, g, b))
        out = "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
        return {"c": out, "a": round(alpha, 3)} if alpha < 1 else {"c": out}


# --------------------------------------------------------------------------- styles
BUILTIN_FMT = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
    12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm", 37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
    46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}


def read_styles(pkg, colors):
    st = pkg.xml("xl/styles.xml")
    fmts = dict(BUILTIN_FMT)
    for nf in st.findall("m:numFmts/m:numFmt", NS):
        fmts[int(nf.get("numFmtId"))] = nf.get("formatCode")

    def font(el):
        if el is None:
            return {}
        f = {}
        for ch in el:
            t = etree.QName(ch).localname
            if t == "b" and ch.get("val", "1") not in ("0", "false"):
                f["b"] = 1
            elif t == "i" and ch.get("val", "1") not in ("0", "false"):
                f["i"] = 1
            elif t == "u":
                f["u"] = ch.get("val", "single")
            elif t == "strike" and ch.get("val", "1") not in ("0", "false"):
                f["s"] = 1
            elif t == "sz":
                f["sz"] = float(ch.get("val"))
            elif t == "name":
                f["name"] = ch.get("val")
            elif t == "color":
                c = colors.sml(ch)
                if c:
                    f["color"] = c
            elif t == "vertAlign":
                f["va"] = ch.get("val")
        return f

    def fill(el):
        if el is None:
            return None
        pf = el.find("m:patternFill", NS)
        if pf is not None:
            pt = pf.get("patternType")
            if pt in (None, "none"):
                # dxf fills often omit patternType but carry bgColor
                if pt is None:
                    c = colors.sml(pf.find("m:bgColor", NS)) or colors.sml(pf.find("m:fgColor", NS))
                    return c
                return None
            if pt == "solid":
                return colors.sml(pf.find("m:fgColor", NS)) or colors.sml(pf.find("m:bgColor", NS))
            return colors.sml(pf.find("m:fgColor", NS)) or "D9D9D9"
        gf = el.find("m:gradientFill", NS)
        if gf is not None:
            stops = gf.findall("m:stop/m:color", NS)
            if stops:
                return colors.sml(stops[0])
        return None

    def border(el):
        if el is None:
            return None
        b = {}
        for side, key in (("left", "l"), ("right", "r"), ("top", "t"), ("bottom", "b")):
            s = el.find("m:" + side, NS)
            if s is not None and s.get("style"):
                b[key] = [s.get("style"), colors.sml(s.find("m:color", NS)) or "000000"]
        return b or None

    fonts = [font(f) for f in st.findall("m:fonts/m:font", NS)]
    fills = [fill(f) for f in st.findall("m:fills/m:fill", NS)]
    borders = [border(b) for b in st.findall("m:borders/m:border", NS)]

    xfs = []
    for xf in st.findall("m:cellXfs/m:xf", NS):
        s = {}
        fid = int(xf.get("numFmtId", 0))
        if fid:
            s["fmt"] = fmts.get(fid, "General")
        fo = fonts[int(xf.get("fontId", 0))]
        if fo:
            s["font"] = fo
        fi = fills[int(xf.get("fillId", 0))]
        if fi:
            s["fill"] = fi
        bo = borders[int(xf.get("borderId", 0))]
        if bo:
            s["border"] = bo
        al = xf.find("m:alignment", NS)
        if al is not None:
            a = {}
            for k, kk in (("horizontal", "h"), ("vertical", "v"), ("indent", "ind"), ("textRotation", "rot")):
                if al.get(k):
                    a[kk] = al.get(k)
            if al.get("wrapText") in ("1", "true"):
                a["wrap"] = 1
            if al.get("shrinkToFit") in ("1", "true"):
                a["shrink"] = 1
            if a:
                s["align"] = a
        pr = xf.find("m:protection", NS)
        if pr is not None and pr.get("locked") in ("0", "false"):
            s["unlocked"] = 1
        if pr is not None and pr.get("hidden") in ("1", "true"):
            s["hidden"] = 1
        xfs.append(s)

    dxfs = []
    for d in st.findall("m:dxfs/m:dxf", NS):
        s = {}
        fo = font(d.find("m:font", NS))
        if fo:
            s["font"] = fo
        fi = fill(d.find("m:fill", NS))
        if fi:
            s["fill"] = fi
        bo = border(d.find("m:border", NS))
        if bo:
            s["border"] = bo
        nf = d.find("m:numFmt", NS)
        if nf is not None:
            s["fmt"] = nf.get("formatCode")
        dxfs.append(s)
    default_font = fonts[0] if fonts else {"name": "Calibri", "sz": 11}
    return xfs, dxfs, default_font


# --------------------------------------------------------------------------- shared strings
def read_shared_strings(pkg):
    out = []
    if not pkg.has("xl/sharedStrings.xml"):
        return out
    root = pkg.xml("xl/sharedStrings.xml")
    for si in root.findall("m:si", NS):
        parts = []
        for t in si.iter(M + "t"):
            # skip phonetic runs
            if etree.QName(t.getparent()).localname == "rPh":
                continue
            parts.append(t.text or "")
        out.append("".join(parts))
    return out


def rich_runs(si, colors):
    """Return rich-text runs for an inline/shared string if it has formatting."""
    runs = []
    for r in si.findall("m:r", NS):
        t = "".join(x.text or "" for x in r.findall("m:t", NS))
        rp = r.find("m:rPr", NS)
        f = {}
        if rp is not None:
            for ch in rp:
                n = etree.QName(ch).localname
                if n == "b":
                    f["b"] = 1
                elif n == "i":
                    f["i"] = 1
                elif n == "u":
                    f["u"] = 1
                elif n == "sz":
                    f["sz"] = float(ch.get("val"))
                elif n == "color":
                    c = colors.sml(ch)
                    if c:
                        f["color"] = c
        runs.append([t, f])
    return runs


# --------------------------------------------------------------------------- drawings
EMU_PER_PX = 9525


def text_body(tx, colors, default_color=None):
    if tx is None:
        return None
    body = tx.find("a:bodyPr", NS)
    out = {"paras": []}
    if body is not None:
        if body.get("anchor"):
            out["anchor"] = body.get("anchor")
        if body.get("vert") and body.get("vert") != "horz":
            out["vert"] = body.get("vert")
        ins = [body.get(k) for k in ("lIns", "tIns", "rIns", "bIns")]
        if any(ins):
            out["ins"] = [int(x) / EMU_PER_PX if x else d for x, d in zip(ins, (9.6, 4.8, 9.6, 4.8))]
        if body.get("wrap") == "none":
            out["nowrap"] = 1
        if body.get("vertOverflow") == "clip":
            out["clip"] = 1
    lst = tx.find("a:lstStyle", NS)
    has_text = False
    for p in tx.findall("a:p", NS):
        para = {"runs": []}
        ppr = p.find("a:pPr", NS)
        if ppr is not None and ppr.get("algn"):
            para["align"] = ppr.get("algn")
        endp = p.find("a:endParaRPr", NS)
        for r in p:
            tag = etree.QName(r).localname
            if tag not in ("r", "fld", "br"):
                continue
            if tag == "br":
                para["runs"].append({"t": "\n"})
                continue
            t = r.find("a:t", NS)
            run = {"t": t.text if t is not None and t.text else ""}
            if run["t"].strip():
                has_text = True
            rpr = r.find("a:rPr", NS)
            if rpr is not None:
                if rpr.get("sz"):
                    run["sz"] = int(rpr.get("sz")) / 100
                if rpr.get("b") == "1":
                    run["b"] = 1
                if rpr.get("i") == "1":
                    run["i"] = 1
                if rpr.get("u") and rpr.get("u") != "none":
                    run["u"] = 1
                sf = rpr.find("a:solidFill", NS)
                col = colors.dml(sf)
                if col and not col.get("ph"):
                    run["color"] = col
                lat = rpr.find("a:latin", NS)
                if lat is not None and not lat.get("typeface", "").startswith("+"):
                    run["font"] = lat.get("typeface")
            if tag == "fld" and r.get("type"):
                run["fld"] = r.get("type")
            para["runs"].append(run)
        if endp is not None and not para["runs"] and endp.get("sz"):
            para["sz"] = int(endp.get("sz")) / 100
        out["paras"].append(para)
    out["hasText"] = has_text
    return out


def sp_props(sppr, style, colors):
    """Fill / line for a shape's spPr, falling back to its style references."""
    res = {}
    style_fill = style_line = style_font = None
    if style is not None:
        fr = style.find("a:fillRef", NS)
        if fr is not None and fr.get("idx") != "0":
            style_fill = colors.dml(fr)
        lr = style.find("a:lnRef", NS)
        if lr is not None and lr.get("idx") != "0":
            style_line = colors.dml(lr)
        fo = style.find("a:fontRef", NS)
        if fo is not None:
            style_font = colors.dml(fo)
    fill_set = line_set = False
    if sppr is not None:
        if sppr.find("a:noFill", NS) is not None:
            res["fill"] = None
            fill_set = True
        sf = sppr.find("a:solidFill", NS)
        if sf is not None:
            res["fill"] = colors.dml(sf)
            fill_set = True
        gf = sppr.find("a:gradFill", NS)
        if gf is not None:
            stops = []
            for gs in gf.findall("a:gsLst/a:gs", NS):
                c = colors.dml(gs)
                if c and not c.get("ph"):
                    stops.append([int(gs.get("pos")) / 1000, c])
            lin = gf.find("a:lin", NS)
            res["grad"] = {"stops": stops, "ang": int(lin.get("ang")) / 60000 if lin is not None and lin.get("ang") else 90}
            fill_set = True
        bf = sppr.find("a:blipFill", NS)
        if bf is not None:
            fill_set = True
        if sppr.find("a:grpFill", NS) is not None:
            res["grpFill"] = 1
            fill_set = True
        ln = sppr.find("a:ln", NS)
        if ln is not None:
            line_set = True
            if ln.find("a:noFill", NS) is not None:
                res["line"] = None
            else:
                lc = colors.dml(ln.find("a:solidFill", NS))
                if lc is None and style_line and not style_line.get("ph"):
                    lc = style_line
                w = int(ln.get("w")) / EMU_PER_PX if ln.get("w") else 0.75
                dash = ln.find("a:prstDash", NS)
                res["line"] = {"c": lc, "w": round(w, 2)}
                if dash is not None and dash.get("val") != "solid":
                    res["line"]["dash"] = dash.get("val")
                he = ln.find("a:headEnd", NS)
                te = ln.find("a:tailEnd", NS)
                if he is not None and he.get("type") not in (None, "none"):
                    res["line"]["head"] = he.get("type")
                if te is not None and te.get("type") not in (None, "none"):
                    res["line"]["tail"] = te.get("type")
        geom = sppr.find("a:prstGeom", NS)
        if geom is not None:
            res["geom"] = geom.get("prst")
            av = {}
            for gd in geom.findall("a:avLst/a:gd", NS):
                fm = gd.get("fmla", "")
                if fm.startswith("val "):
                    av[gd.get("name")] = int(fm[4:])
            if av:
                res["av"] = av
        elif sppr.find("a:custGeom", NS) is not None:
            res["geom"] = "custom"
        xfrm = sppr.find("a:xfrm", NS)
        if xfrm is not None:
            if xfrm.get("rot"):
                res["rot"] = int(xfrm.get("rot")) / 60000
            if xfrm.get("flipH") == "1":
                res["flipH"] = 1
            if xfrm.get("flipV") == "1":
                res["flipV"] = 1
        ef = sppr.find("a:effectLst/a:outerShdw", NS)
        if ef is not None:
            res["shadow"] = 1
    if not fill_set and style_fill is not None and not style_fill.get("ph"):
        res["fill"] = style_fill
    if not line_set and style_line is not None and not style_line.get("ph"):
        res["line"] = {"c": style_line, "w": 1}
    if style_font is not None and not style_font.get("ph"):
        res["fontColor"] = style_font
    return res


def xfrm_box(sppr):
    if sppr is None:
        return None
    x = sppr.find("a:xfrm", NS)
    if x is None:
        return None
    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return [int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))]


def parse_obj(el, colors, rels, media_out, chart_ids):
    tag = etree.QName(el).localname
    nv = el.find(".//xdr:cNvPr", NS)
    o = {"kind": tag}
    if nv is not None:
        o["name"] = nv.get("name")
        o["id"] = nv.get("id")
        if nv.get("descr"):
            o["descr"] = nv.get("descr")
        if nv.get("hidden") == "1":
            o["hidden"] = 1
        hl = nv.find("a:hlinkClick", NS)
        if hl is not None and hl.get(R_ID) and hl.get(R_ID) in rels:
            o["href"] = rels[hl.get(R_ID)]["target"]
    if el.get("macro"):
        o["macro"] = el.get("macro").replace("[0]!", "")
    if tag == "sp" or tag == "cxnSp":
        sppr = el.find("xdr:spPr", NS)
        o.update(sp_props(sppr, el.find("xdr:style", NS), colors))
        o["box"] = xfrm_box(sppr)
        tb = text_body(el.find("xdr:txBody", NS), colors)
        if tb and tb["paras"]:
            o["text"] = tb
        # textlink (=A1)
        if el.get("textlink"):
            o["textlink"] = el.get("textlink")
    elif tag == "pic":
        sppr = el.find("xdr:spPr", NS)
        o.update(sp_props(sppr, None, colors))
        o["box"] = xfrm_box(sppr)
        blip = el.find(".//a:blip", NS)
        if blip is not None:
            rid = blip.get(R_EMBED)
            # prefer the svg extension if any
            svg = blip.find(".//{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip")
            if svg is not None and svg.get(R_EMBED) in rels:
                rid = svg.get(R_EMBED)
            if rid in rels:
                tgt = rels[rid]["target"]
                media_out.add(tgt)
                o["img"] = "media/" + posixpath.basename(tgt)
        src = el.find(".//a:srcRect", NS)
        if src is not None and any(src.get(k) for k in ("l", "t", "r", "b")):
            o["crop"] = [int(src.get(k, "0")) / 100000 for k in ("l", "t", "r", "b")]
    elif tag == "graphicFrame":
        ch = el.find(".//c:chart", NS)
        if ch is not None and ch.get(R_ID) in rels:
            tgt = rels[ch.get(R_ID)]["target"]
            o["chart"] = posixpath.basename(tgt).replace(".xml", "")
            chart_ids.add(tgt)
        x = el.find("xdr:xfrm", NS)
        if x is not None:
            off, ext = x.find("a:off", NS), x.find("a:ext", NS)
            o["box"] = [int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))]
    elif tag == "grpSp":
        gp = el.find("xdr:grpSpPr", NS)
        x = gp.find("a:xfrm", NS) if gp is not None else None
        if x is not None:
            off, ext = x.find("a:off", NS), x.find("a:ext", NS)
            choff, chext = x.find("a:chOff", NS), x.find("a:chExt", NS)
            o["box"] = [int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))]
            o["chbox"] = [int(choff.get("x")), int(choff.get("y")), int(chext.get("cx")), int(chext.get("cy"))]
        if gp is not None:
            gprops = sp_props(gp, None, colors)
            for k in ("fill", "grad"):
                if gprops.get(k):
                    o[k] = gprops[k]
        o["children"] = []
        for ch in el:
            ct = etree.QName(ch).localname
            if ct in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                c = parse_obj(ch, colors, rels, media_out, chart_ids)
                if c.pop("grpFill", None):
                    for k in ("fill", "grad"):
                        if o.get(k):
                            c[k] = o[k]
                o["children"].append(c)
    return o


def parse_drawing(pkg, part, colors, media_out, chart_ids):
    root = pkg.xml(part)
    rels = pkg.rels(part)
    out = []
    for anc in root:
        at = etree.QName(anc).localname
        if at not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            # mc:AlternateContent wrapping
            if at == "AlternateContent":
                ch = anc.find("mc:Choice", NS)
                if ch is not None and len(ch):
                    anc = ch[0]
                    at = etree.QName(anc).localname
                else:
                    continue
            else:
                continue
        a = {"anchor": at}
        if anc.get("editAs"):
            a["editAs"] = anc.get("editAs")

        def pos(e):
            return [int(e.find("xdr:col", NS).text), int(e.find("xdr:colOff", NS).text),
                    int(e.find("xdr:row", NS).text), int(e.find("xdr:rowOff", NS).text)]

        fr, to = anc.find("xdr:from", NS), anc.find("xdr:to", NS)
        if fr is not None:
            a["from"] = pos(fr)
        if to is not None:
            a["to"] = pos(to)
        ext = anc.find("xdr:ext", NS)
        if ext is not None:
            a["ext"] = [int(ext.get("cx")), int(ext.get("cy"))]
        p = anc.find("xdr:pos", NS)
        if p is not None:
            a["pos"] = [int(p.get("x")), int(p.get("y"))]
        for ch in anc:
            ct = etree.QName(ch).localname
            if ct == "AlternateContent":
                c2 = ch.find("mc:Choice", NS)
                if c2 is None or not len(c2):
                    c2 = ch.find("mc:Fallback", NS)
                if c2 is not None and len(c2):
                    ch = c2[0]
                    ct = etree.QName(ch).localname
            if ct in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                a["obj"] = parse_obj(ch, colors, rels, media_out, chart_ids)
        if "obj" in a:
            out.append(a)
    return out


# --------------------------------------------------------------------------- charts
def chart_text(el):
    if el is None:
        return None
    rich = el.find(".//c:rich", NS)
    if rich is not None:
        lines = []
        for p in rich.findall("a:p", NS):
            lines.append("".join(t.text or "" for t in p.iter("{%s}t" % NS["a"])))
        return {"text": "\n".join(lines)}
    sr = el.find(".//c:strRef", NS)
    if sr is not None:
        cache = [pt.findtext("c:v", namespaces=NS) for pt in sr.findall("c:strCache/c:pt", NS)]
        return {"ref": sr.findtext("c:f", namespaces=NS), "cache": " ".join(c or "" for c in cache)}
    return None


def chart_data(el):
    """c:cat / c:val / c:xVal / c:yVal / c:tx  -> {ref, cache, fmt}"""
    if el is None:
        return None
    for kind in ("numRef", "strRef", "multiLvlStrRef"):
        r = el.find("c:" + kind, NS)
        if r is not None:
            out = {"ref": r.findtext("c:f", namespaces=NS)}
            cache_el = r.find("c:numCache", NS) if kind == "numRef" else r.find("c:strCache", NS)
            if kind == "multiLvlStrRef":
                cache_el = r.find("c:multiLvlStrCache/c:lvl", NS)
            if cache_el is not None:
                n = int(cache_el.findtext("c:ptCount", default="0", namespaces=NS) or 0)
                vals = [None] * n
                for pt in cache_el.findall("c:pt", NS):
                    i = int(pt.get("idx"))
                    v = pt.findtext("c:v", namespaces=NS)
                    if i < n:
                        if kind == "numRef":
                            try:
                                vals[i] = float(v)
                            except (TypeError, ValueError):
                                vals[i] = None
                        else:
                            vals[i] = v
                out["cache"] = vals
                fc = cache_el.findtext("c:formatCode", namespaces=NS)
                if fc:
                    out["fmt"] = fc
            return out
    for kind in ("numLit", "strLit"):
        r = el.find("c:" + kind, NS)
        if r is not None:
            n = int(r.findtext("c:ptCount", default="0", namespaces=NS) or 0)
            pts = r.findall("c:pt", NS)
            n = max([n] + [int(pt.get("idx")) + 1 for pt in pts])
            vals = [None] * n
            for pt in pts:
                v = pt.findtext("c:v", namespaces=NS)
                vals[int(pt.get("idx"))] = float(v) if kind == "numLit" else v
            return {"lit": vals}
    v = el.findtext("c:v", namespaces=NS)
    if v is not None:
        return {"lit": [v]}
    return None


def chart_sppr(el, colors):
    if el is None:
        return {}
    sppr = el.find("c:spPr", NS)
    out = {}
    if sppr is None:
        return out
    if sppr.find("a:noFill", NS) is not None:
        out["fill"] = None
    sf = sppr.find("a:solidFill", NS)
    if sf is not None:
        out["fill"] = colors.dml(sf)
    gf = sppr.find("a:gradFill", NS)
    if gf is not None:
        gs = gf.find("a:gsLst/a:gs", NS)
        if gs is not None:
            out["fill"] = colors.dml(gs)
    ln = sppr.find("a:ln", NS)
    if ln is not None:
        if ln.find("a:noFill", NS) is not None:
            out["line"] = None
        else:
            lc = colors.dml(ln.find("a:solidFill", NS))
            out["line"] = {"c": lc, "w": round(int(ln.get("w")) / 12700, 2) if ln.get("w") else None}
            d = ln.find("a:prstDash", NS)
            if d is not None and d.get("val") != "solid":
                out["line"]["dash"] = d.get("val")
    return out


def txpr_font(el, colors):
    if el is None:
        return None
    tx = el.find("c:txPr", NS)
    if tx is None:
        return None
    d = tx.find(".//a:defRPr", NS)
    if d is None:
        return None
    f = {}
    if d.get("sz"):
        f["sz"] = int(d.get("sz")) / 100
    if d.get("b") == "1":
        f["b"] = 1
    c = colors.dml(d.find("a:solidFill", NS))
    if c and not c.get("ph"):
        f["color"] = c
    return f or None


def parse_dlbls(el, colors):
    if el is None:
        return None
    d = {}
    for k in ("showVal", "showPercent", "showCatName", "showSerName", "showLegendKey"):
        e = el.find("c:" + k, NS)
        if e is not None and e.get("val") == "1":
            d[k] = 1
    if el.find("c:delete", NS) is not None and el.find("c:delete", NS).get("val") == "1":
        return {"delete": 1}
    p = el.find("c:dLblPos", NS)
    if p is not None:
        d["pos"] = p.get("val")
    nf = el.find("c:numFmt", NS)
    if nf is not None:
        d["fmt"] = nf.get("formatCode")
    f = txpr_font(el, colors)
    if f:
        d["font"] = f
    return d or None


def parse_chart(pkg, part, colors):
    root = pkg.xml(part)
    ch = root.find("c:chart", NS)
    out = {"plots": [], "axes": {}}
    t = ch.find("c:title", NS)
    if t is not None:
        out["title"] = chart_text(t.find("c:tx", NS)) or {"auto": 1}
        f = txpr_font(t, colors) or None
        if t.find(".//a:defRPr", NS) is not None and not f:
            d = t.find(".//a:defRPr", NS)
            f = {"sz": int(d.get("sz")) / 100} if d.get("sz") else None
        if f:
            out["titleFont"] = f
        ov = t.find("c:overlay", NS)
    atd = ch.find("c:autoTitleDeleted", NS)
    if atd is not None and atd.get("val") == "1":
        out["autoTitleDeleted"] = 1
    pa = ch.find("c:plotArea", NS)
    for p in pa:
        tag = etree.QName(p).localname
        if not tag.endswith("Chart"):
            continue
        plot = {"type": tag.replace("3D", "").replace("Chart", ""), "series": []}
        for k in ("barDir", "grouping", "varyColors", "gapWidth", "overlap", "holeSize", "scatterStyle", "firstSliceAng"):
            e = p.find("c:" + k, NS)
            if e is not None:
                plot[k] = e.get("val")
        plot["axIds"] = [e.get("val") for e in p.findall("c:axId", NS)]
        dl = parse_dlbls(p.find("c:dLbls", NS), colors)
        if dl:
            plot["dLbls"] = dl
        for s in p.findall("c:ser", NS):
            ser = {"idx": int(s.find("c:idx", NS).get("val")), "order": int(s.find("c:order", NS).get("val"))}
            tx = s.find("c:tx", NS)
            if tx is not None:
                ser["tx"] = chart_data(tx)
            for k in ("cat", "val", "xVal", "yVal"):
                d = chart_data(s.find("c:" + k, NS))
                if d:
                    ser[k] = d
            ser.update(chart_sppr(s, colors))
            mk = s.find("c:marker", NS)
            if mk is not None:
                sym = mk.find("c:symbol", NS)
                ser["marker"] = {"symbol": sym.get("val") if sym is not None else None}
                sz = mk.find("c:size", NS)
                if sz is not None:
                    ser["marker"]["size"] = int(sz.get("val"))
                ser["marker"].update(chart_sppr(mk, colors))
            sm = s.find("c:smooth", NS)
            if sm is not None and sm.get("val") == "1":
                ser["smooth"] = 1
            dpts = []
            for dp in s.findall("c:dPt", NS):
                x = {"idx": int(dp.find("c:idx", NS).get("val"))}
                x.update(chart_sppr(dp, colors))
                dpts.append(x)
            if dpts:
                ser["dPt"] = dpts
            dl = parse_dlbls(s.find("c:dLbls", NS), colors)
            if dl:
                ser["dLbls"] = dl
            tl = s.find("c:trendline", NS)
            if tl is not None:
                ser["trendline"] = tl.find("c:trendlineType", NS).get("val")
            plot["series"].append(ser)
        out["plots"].append(plot)
    for ax in pa:
        tag = etree.QName(ax).localname
        if tag not in ("catAx", "valAx", "dateAx", "serAx"):
            continue
        a = {"type": tag}
        for k in ("axPos", "delete", "crosses", "crossAx", "tickLblPos", "majorTickMark"):
            e = ax.find("c:" + k, NS)
            if e is not None:
                a[k] = e.get("val")
        sc = ax.find("c:scaling", NS)
        if sc is not None:
            for k in ("min", "max", "orientation", "logBase"):
                e = sc.find("c:" + k, NS)
                if e is not None:
                    a[k] = e.get("val")
        nf = ax.find("c:numFmt", NS)
        if nf is not None:
            a["fmt"] = nf.get("formatCode")
        if ax.find("c:majorGridlines", NS) is not None:
            a["grid"] = 1
        tt = ax.find("c:title", NS)
        if tt is not None:
            a["title"] = chart_text(tt.find("c:tx", NS))
        f = txpr_font(ax, colors)
        if f:
            a["font"] = f
        a.update({"line": chart_sppr(ax, colors).get("line", "unset")} if ax.find("c:spPr", NS) is not None else {})
        out["axes"][ax.find("c:axId", NS).get("val")] = a
    lg = ch.find("c:legend", NS)
    if lg is not None:
        lp = lg.find("c:legendPos", NS)
        out["legend"] = {"pos": lp.get("val") if lp is not None else "r"}
        deleted = [int(le.find("c:idx", NS).get("val")) for le in lg.findall("c:legendEntry", NS)
                   if le.find("c:delete", NS) is not None and le.find("c:delete", NS).get("val") == "1"]
        if deleted:
            out["legend"]["deleted"] = deleted
        f = txpr_font(lg, colors)
        if f:
            out["legend"]["font"] = f
    db = ch.find("c:dispBlanksAs", NS)
    if db is not None:
        out["dispBlanksAs"] = db.get("val")
    out["chartSpace"] = chart_sppr(root, colors)
    out["plotAreaFill"] = chart_sppr(pa, colors)
    rc = root.find("c:roundedCorners", NS)
    out["rounded"] = 1 if (rc is None or rc.get("val") == "1") else 0
    piv = root.find("c:pivotSource", NS)
    if piv is not None:
        out["pivot"] = piv.findtext("c:name", namespaces=NS)
    return out


# --------------------------------------------------------------------------- formulas / cells
def read_sheet(pkg, part, idx, sst, colors, xfs, media_out, chart_ids, sheet_meta):
    rels = pkg.rels(part)
    sheet = {"rows": {}, "cols": [], "merges": [], "cells": [], "formulas": [], "shared": {},
             "dv": [], "cf": [], "links": [], "drawing": [], "controls": [], "comments": []}
    rich_cells = []
    ctx = etree.iterparse(pkg.z.open(part), events=("start", "end"))
    in_sheetdata = False
    for ev, el in ctx:
        tag = etree.QName(el).localname
        if ev == "start":
            if tag == "sheetData":
                in_sheetdata = True
            continue
        if tag == "sheetPr":
            tc = el.find("m:tabColor", NS)
            if tc is not None:
                sheet["tabColor"] = colors.sml(tc)
            if el.get("codeName"):
                sheet["codeName"] = el.get("codeName")
        elif tag == "sheetView":
            sheet["showGrid"] = el.get("showGridLines") != "0"
            if el.get("zoomScale"):
                sheet["zoom"] = int(el.get("zoomScale"))
            pane = el.find("m:pane", NS)
            if pane is not None and pane.get("state") in ("frozen", "frozenSplit"):
                sheet["freeze"] = [int(float(pane.get("ySplit", "0"))), int(float(pane.get("xSplit", "0")))]
            if el.get("topLeftCell"):
                sheet["topLeft"] = el.get("topLeftCell")
            if el.get("showRowColHeaders") == "0":
                sheet["noHeadings"] = 1
        elif tag == "sheetFormatPr":
            sheet["defRowH"] = float(el.get("defaultRowHeight", "15"))
            if el.get("defaultColWidth"):
                sheet["defColW"] = float(el.get("defaultColWidth"))
            sheet["baseColW"] = int(el.get("baseColWidth", "8"))
            if el.get("zeroHeight") == "1":
                sheet["zeroHeight"] = 1
        elif tag == "col" and not in_sheetdata and etree.QName(el).namespace == NS["m"]:
            c = [int(el.get("min")), int(el.get("max")), float(el.get("width", "9.140625")),
                 1 if el.get("hidden") in ("1", "true") else 0, int(el.get("style", "0"))]
            sheet["cols"].append(c)
        elif tag == "row" and etree.QName(el).namespace == NS["m"]:
            r = int(el.get("r"))
            info = {}
            if el.get("ht") and el.get("customHeight") in ("1", "true"):
                info["h"] = float(el.get("ht"))
            elif el.get("ht"):
                info["h"] = float(el.get("ht"))
            if el.get("hidden") in ("1", "true"):
                info["hidden"] = 1
            if el.get("s") and el.get("customFormat") in ("1", "true"):
                info["s"] = int(el.get("s"))
            if info:
                sheet["rows"][r] = info
            for c in el.findall("m:c", NS):
                ref = c.get("r")
                rr, cc = parse_cell(ref)
                s = int(c.get("s", "0"))
                t = c.get("t", "n")
                v = c.find("m:v", NS)
                f = c.find("m:f", NS)
                val = None
                if t == "s" and v is not None:
                    val = sst[int(v.text)]
                elif t == "inlineStr":
                    is_ = c.find("m:is", NS)
                    val = "".join(x.text or "" for x in is_.iter(M + "t")) if is_ is not None else ""
                elif t == "b" and v is not None:
                    val = v.text in ("1", "true")
                elif t == "e" and v is not None:
                    val = {"e": v.text}
                elif t == "str" and v is not None:
                    val = v.text or ""
                elif t == "str" and v is None:
                    val = ""
                elif v is not None and v.text is not None:
                    try:
                        x = float(v.text)
                        val = int(x) if x.is_integer() and abs(x) < 2 ** 53 else x
                    except ValueError:
                        val = v.text
                if c.get("vm"):
                    rich_cells.append([rr, cc, c.get("vm")])
                if f is not None:
                    ft = f.get("t")
                    rec = [rr, cc]
                    text = f.text
                    kind = ""
                    if ft == "shared":
                        si = int(f.get("si"))
                        if text:
                            sheet["shared"][si] = [rr, cc, text]
                            rec += [text, si]
                        else:
                            rec += [None, si]
                    elif ft == "array":
                        kind = "d" if c.get("cm") else "a"
                        rec += [text or "", None, kind, f.get("ref") or ref]
                    elif ft == "dataTable":
                        rec += [None, None, "t", f.get("ref")]
                    else:
                        rec += [text or "", None]
                    sheet["formulas"].append(rec)
                if val is None and s == 0 and f is None:
                    continue
                cell = [rr, cc, val]
                if s:
                    cell.append(s)
                sheet["cells"].append(cell)
            el.clear()
        elif tag == "sheetData":
            in_sheetdata = False
        elif tag == "mergeCell":
            sheet["merges"].append(parse_range(el.get("ref")))
        elif tag == "sheetProtection":
            if el.get("sheet") in ("1", "true"):
                sheet["protected"] = 1
        elif tag == "dataValidation" and etree.QName(el).namespace == NS["m"]:
            sheet["dv"].append(dv_rec(el, el.get("sqref"), el.findtext("m:formula1", namespaces=NS),
                                      el.findtext("m:formula2", namespaces=NS)))
        elif tag == "dataValidation" and etree.QName(el).namespace == NS["x14"]:
            sheet["dv"].append(dv_rec(el, el.findtext("xm:sqref", namespaces=NS),
                                      el.findtext("x14:formula1/xm:f", namespaces=NS),
                                      el.findtext("x14:formula2/xm:f", namespaces=NS)))
        elif tag == "conditionalFormatting" and etree.QName(el).namespace == NS["m"]:
            rules = []
            for rule in el.findall("m:cfRule", NS):
                rr_ = {"type": rule.get("type"), "pri": int(rule.get("priority", "0"))}
                for k in ("dxfId", "operator", "text", "stopIfTrue", "timePeriod", "rank", "percent", "bottom", "aboveAverage"):
                    if rule.get(k) is not None:
                        rr_[k] = rule.get(k)
                rr_["f"] = [x.text for x in rule.findall("m:formula", NS)]
                cs = rule.find("m:colorScale", NS)
                if cs is not None:
                    rr_["scale"] = {"cfvo": [[x.get("type"), x.get("val")] for x in cs.findall("m:cfvo", NS)],
                                    "colors": [colors.sml(x) for x in cs.findall("m:color", NS)]}
                db = rule.find("m:dataBar", NS)
                if db is not None:
                    rr_["bar"] = colors.sml(db.find("m:color", NS))
                rules.append(rr_)
            sheet["cf"].append({"sqref": sqref_list(el.get("sqref")), "rules": rules})
        elif tag == "hyperlink" and etree.QName(el).namespace == NS["m"]:
            h = {"ref": parse_range(el.get("ref"))}
            if el.get(R_ID) and el.get(R_ID) in rels:
                h["url"] = rels[el.get(R_ID)]["target"]
            if el.get("location"):
                h["loc"] = el.get("location")
            if el.get("display"):
                h["display"] = el.get("display")
            sheet["links"].append(h)
        elif tag == "drawing" and etree.QName(el).namespace == NS["m"]:
            rid = el.get(R_ID)
            if rid in rels:
                sheet["drawing"] = parse_drawing(pkg, rels[rid]["target"], colors, media_out, chart_ids)
        elif tag == "control" and etree.QName(el).namespace == NS["m"]:
            rid = el.get(R_ID)
            ctl = {"name": el.get("name"), "shapeId": el.get("shapeId")}
            pr = el.find("m:controlPr", NS)
            if pr is not None:
                if pr.get("macro"):
                    ctl["macro"] = pr.get("macro").replace("[0]!", "")
                anc = pr.find("m:anchor", NS)
                if anc is not None:
                    fr = anc.find("m:from", NS)
                    to = anc.find("m:to", NS)

                    def p2(e):
                        return [int(e.findtext("xdr:col", namespaces=NS)), int(e.findtext("xdr:colOff", namespaces=NS)),
                                int(e.findtext("xdr:row", namespaces=NS)), int(e.findtext("xdr:rowOff", namespaces=NS))]

                    ctl["from"] = p2(fr)
                    ctl["to"] = p2(to)
            if rid in rels and pkg.has(rels[rid]["target"]):
                cp = pkg.xml(rels[rid]["target"])
                for k in ("objectType", "fmlaLink", "checked", "fmlaRange", "sel", "fmlaGroup", "noThreeD", "dropLines", "min", "max", "inc", "page", "val"):
                    if cp.get(k) is not None:
                        ctl[k] = cp.get(k)
            sheet["controls"].append(ctl)
    # comments
    for rid, rel in rels.items():
        if rel["type"] == "comments" and pkg.has(rel["target"]):
            croot = pkg.xml(rel["target"])
            authors = [a.text for a in croot.findall("m:authors/m:author", NS)]
            for cm in croot.findall("m:commentList/m:comment", NS):
                txt = "".join(t.text or "" for t in cm.iter(M + "t"))
                ai = int(cm.get("authorId", "0"))
                r_, c_ = parse_cell(cm.get("ref"))
                sheet["comments"].append([r_, c_, txt.strip(), authors[ai] if ai < len(authors) else ""])
        if rel["type"] == "table":
            sheet_meta.setdefault("tables", []).append(rel["target"])
    if rich_cells:
        sheet["rich"] = rich_cells
    return sheet


def dv_rec(el, sqref, f1, f2):
    d = {"sqref": sqref_list(sqref), "type": el.get("type", "any")}
    for k in ("operator", "allowBlank", "showDropDown", "showErrorMessage", "showInputMessage", "prompt", "promptTitle", "error", "errorTitle", "errorStyle"):
        if el.get(k) is not None:
            d[k] = el.get(k)
    if f1 is not None:
        d["f1"] = f1
    if f2 is not None:
        d["f2"] = f2
    return d


def read_tables(pkg, parts, sheet_index):
    out = []
    for part, sname in parts:
        t = pkg.xml(part)
        cols = [c.get("name") for c in t.findall("m:tableColumns/m:tableColumn", NS)]
        calc = {}
        for c in t.findall("m:tableColumns/m:tableColumn", NS):
            cf = c.find("m:calculatedColumnFormula", NS)
            if cf is not None:
                calc[c.get("name")] = cf.text
        out.append({"name": t.get("displayName") or t.get("name"), "sheet": sname, "ref": parse_range(t.get("ref")),
                    "header": int(t.get("headerRowCount", "1")), "totals": int(t.get("totalsRowCount", "0")),
                    "columns": cols, "style": (t.find("m:tableStyleInfo", NS).get("name") if t.find("m:tableStyleInfo", NS) is not None else None),
                    "stripes": (t.find("m:tableStyleInfo", NS).get("showRowStripes") == "1" if t.find("m:tableStyleInfo", NS) is not None else False)})
    return out


def read_external(pkg):
    out = {}
    wbrels = pkg.rels("xl/workbook.xml")
    wb = pkg.xml("xl/workbook.xml")
    refs = wb.findall("m:externalReferences/m:externalReference", NS)
    for i, er in enumerate(refs, start=1):
        rel = wbrels.get(er.get(R_ID))
        if not rel or not pkg.has(rel["target"]):
            continue
        x = pkg.xml(rel["target"])
        eb = x.find("m:externalBook", NS)
        names = [s.get("val") for s in eb.findall("m:sheetNames/m:sheetName", NS)]
        data = {}
        for sd in eb.findall("m:sheetDataSet/m:sheetData", NS):
            sid = int(sd.get("sheetId"))
            cells = {}
            for row in sd.findall("m:row", NS):
                for c in row.findall("m:cell", NS):
                    v = c.findtext("m:v", namespaces=NS)
                    t = c.get("t", "n")
                    if v is None:
                        continue
                    if t == "n":
                        x_ = float(v)
                        v = int(x_) if x_.is_integer() else x_
                    elif t == "b":
                        v = v == "1"
                    elif t == "e":
                        v = {"e": v}
                    cells[c.get("r")] = v
            data[names[sid] if sid < len(names) else str(sid)] = cells
        out[str(i)] = {"sheets": names, "data": data}
    return out


def read_pivots(pkg, sheets):
    """Pivot tables (row fields + one summed data field) so the app can refresh them."""
    out = []
    for sh in sheets:
        for rel in pkg.rels(sh["part"]).values():
            if rel["type"] != "pivotTable":
                continue
            pt = pkg.xml(rel["target"])
            cache_part = next((r["target"] for r in pkg.rels(rel["target"]).values() if r["type"] == "pivotCacheDefinition"), None)
            if not cache_part:
                continue
            cd = pkg.xml(cache_part)
            ws = cd.find("m:cacheSource/m:worksheetSource", NS)
            if ws is None or not ws.get("ref"):
                continue
            fields = []
            for cf in cd.findall("m:cacheFields/m:cacheField", NS):
                items = []
                si = cf.find("m:sharedItems", NS)
                if si is not None:
                    for it in si:
                        tag = etree.QName(it).localname
                        items.append(None if tag == "m" else it.get("v"))
                fields.append({"name": cf.get("name"), "items": items})
            pfs = pt.findall("m:pivotFields/m:pivotField", NS)
            rows = [int(f.get("x")) for f in pt.findall("m:rowFields/m:field", NS) if int(f.get("x")) >= 0]
            hidden = {}
            for fi in rows:
                if fi < len(pfs):
                    hs = [int(i.get("x")) for i in pfs[fi].findall("m:items/m:item", NS) if i.get("h") == "1" and i.get("x") is not None]
                    if hs:
                        hidden[fi] = [fields[fi]["items"][x] if x < len(fields[fi]["items"]) else None for x in hs]
            df = pt.find("m:dataFields/m:dataField", NS)
            loc = pt.find("m:location", NS)
            out.append({
                "name": pt.get("name"), "sheet": sh["name"], "ref": loc.get("ref"),
                "firstDataRow": int(loc.get("firstDataRow", "1")),
                "rows": rows, "data": int(df.get("fld")) if df is not None else None,
                "hidden": {str(k): v for k, v in hidden.items()},
                "grandTotal": pt.get("rowGrandTotals") != "0",
                "src": {"sheet": ws.get("sheet") or sh["name"], "ref": ws.get("ref")},
                "fields": [f["name"] for f in fields],
            })
    return out


def main(src, dst):
    pkg = Pkg(src)
    colors = Colors(pkg)
    xfs, dxfs, default_font = read_styles(pkg, colors)
    sst = read_shared_strings(pkg)
    wb = pkg.xml("xl/workbook.xml")
    wbrels = pkg.rels("xl/workbook.xml")
    calc = wb.find("m:calcPr", NS)
    sheets = []
    for i, s in enumerate(wb.findall("m:sheets/m:sheet", NS)):
        sheets.append({"name": s.get("name"), "state": s.get("state", "visible"), "part": wbrels[s.get(R_ID)]["target"], "index": i})
    names = []
    for dn in wb.findall("m:definedNames/m:definedName", NS):
        nm = dn.get("name")
        if nm.startswith("_xlnm._FilterDatabase") or nm.startswith("_xleta") or nm.startswith("_xlpm"):
            continue
        rec = {"name": nm, "f": dn.text}
        if dn.get("localSheetId") is not None:
            rec["sheet"] = int(dn.get("localSheetId"))
        if dn.get("hidden") == "1":
            rec["hidden"] = 1
        names.append(rec)

    os.makedirs(dst, exist_ok=True)
    os.makedirs(os.path.join(dst, "media"), exist_ok=True)
    media_out, chart_ids = set(), set()
    table_parts = []
    meta = []
    for s in sheets:
        sm = {}
        data = read_sheet(pkg, s["part"], s["index"], sst, colors, xfs, media_out, chart_ids, sm)
        for tp in sm.get("tables", []):
            table_parts.append((tp, s["name"]))
        info = {"name": s["name"], "state": s["state"], "file": "sheet%d.json" % s["index"]}
        for k in ("tabColor", "codeName", "protected", "zoom", "showGrid", "freeze", "topLeft"):
            if k in data:
                info[k] = data[k]
        meta.append(info)
        data["name"] = s["name"]
        with open(os.path.join(dst, info["file"]), "w", encoding="utf-8") as fh:
            json.dump(data, fh, separators=(",", ":"), ensure_ascii=False)
        print("sheet %-34s cells=%-7d formulas=%-6d shapes=%d" % (s["name"], len(data["cells"]), len(data["formulas"]), len(data["drawing"])))

    charts = {}
    for cp in sorted(chart_ids):
        charts[posixpath.basename(cp).replace(".xml", "")] = parse_chart(pkg, cp, colors)
    for mp in media_out:
        with open(os.path.join(dst, "media", posixpath.basename(mp)), "wb") as fh:
            fh.write(pkg.raw(mp))

    # rich values (linked data types such as currency pairs) -> display text + fields
    rich = []
    if pkg.has("xl/richData/rdrichvalue.xml"):
        rv = pkg.xml("xl/richData/rdrichvalue.xml")
        structs = pkg.xml("xl/richData/rdrichvaluestructure.xml")
        skeys = [[k.get("n") for k in s_.findall("{*}k")] for s_ in structs.findall("{*}s")]
        for r in rv.findall("{*}rv"):
            keys = skeys[int(r.get("s"))]
            vals = [v.text for v in r.findall("{*}v")]
            rich.append(dict(zip(keys, vals)))
    vmeta = []
    if pkg.has("xl/metadata.xml"):
        md = pkg.xml("xl/metadata.xml")
        rvb = [int(x.get("i")) for x in md.iter("{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}rvb")]
        vmeta = rvb

    workbook = {
        "source": os.path.basename(src),
        "sheets": meta,
        "names": names,
        "tables": read_tables(pkg, table_parts, None),
        "xfs": xfs,
        "dxfs": dxfs,
        "defaultFont": default_font,
        "theme": colors.theme,
        "external": read_external(pkg),
        "charts": charts,
        "pivots": read_pivots(pkg, sheets),
        "rich": rich,
        "richMap": vmeta,
        "calc": {"iterate": calc.get("iterate") == "1" if calc is not None else False,
                 "iterateCount": int(calc.get("iterateCount", "100")) if calc is not None else 100},
    }
    with open(os.path.join(dst, "workbook.json"), "w", encoding="utf-8") as fh:
        json.dump(workbook, fh, separators=(",", ":"), ensure_ascii=False)
    print("charts=%d media=%d tables=%d names=%d" % (len(charts), len(media_out), len(workbook["tables"]), len(names)))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
