"""Page setup (size, margins, columns, header/footer, watermark, borders ...) and
its .docx round trip.  Only python-docx here, no Qt, so it is unit-testable.
"""
import re
from dataclasses import dataclass, replace
from xml.sax.saxutils import quoteattr

from docx.enum.section import WD_ORIENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Mm

PAGE_SIZES = {  # portrait width x height, mm
    "A3": (297.0, 420.0), "A4": (210.0, 297.0), "A5": (148.0, 210.0), "B5": (176.0, 250.0),
    "Letter": (215.9, 279.4), "Legal": (215.9, 355.6), "Executive": (184.2, 266.7),
    "Tabloid": (279.4, 431.8),
}
MARGIN_PRESETS = {  # top, bottom, left, right (mm)
    "Normal": (25.4, 25.4, 25.4, 25.4),
    "Narrow": (12.7, 12.7, 12.7, 12.7),
    "Moderate": (25.4, 25.4, 19.1, 19.1),
    "Wide": (25.4, 25.4, 50.8, 50.8),
    "Mirrored": (25.4, 25.4, 31.8, 25.4),
}
BORDER_STYLES = {"Single": "single", "Double": "double", "Dashed": "dashed", "Dotted": "dotted",
                 "Thick": "thick"}
V_NS = "urn:schemas-microsoft-com:vml"
HF_TOKEN = re.compile(r"(\{pages?\})")


@dataclass
class PageSetup:
    width_mm: float = 210.0       # portrait dimensions
    height_mm: float = 297.0
    orientation: str = "portrait"
    top: float = 25.4
    bottom: float = 25.4
    left: float = 25.4
    right: float = 25.4
    columns: int = 1
    col_gap: float = 12.7
    header: str = ""
    footer: str = ""              # {page} and {pages} become Word fields
    watermark: str = ""
    border: bool = False
    border_style: str = "single"
    border_width: float = 1.0     # pt
    border_color: str = "auto"    # "auto" or RRGGBB
    line_numbers: bool = False
    hyphenate: bool = False

    @property
    def eff_w(self):
        return self.height_mm if self.orientation == "landscape" else self.width_mm

    @property
    def eff_h(self):
        return self.width_mm if self.orientation == "landscape" else self.height_mm

    @property
    def size_name(self):
        for name, (w, h) in PAGE_SIZES.items():
            if abs(w - self.width_mm) < 1 and abs(h - self.height_mm) < 1:
                return name
        return "Custom"

    @property
    def margin_name(self):
        cur = (self.top, self.bottom, self.left, self.right)
        for name, m in MARGIN_PRESETS.items():
            if all(abs(a - b) < 0.2 for a, b in zip(cur, m)):
                return name
        return "Custom"

    def with_(self, **kw):
        return replace(self, **kw)


# --------------------------------------------------------------------------- write
def _set_cols(sectPr, n, gap_mm):
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = parse_xml(f'<w:cols {nsdecls("w")}/>')
        anchor = sectPr.find(qn("w:docGrid"))
        anchor.addprevious(cols) if anchor is not None else sectPr.append(cols)
    for k in list(cols.attrib):
        del cols.attrib[k]
    cols.set(qn("w:space"), str(round(gap_mm * 56.7)))
    if n > 1:
        cols.set(qn("w:num"), str(n))


def _fld(par, code):
    for xml in (f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>',
                f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> {code} </w:instrText></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>',
                f'<w:r {nsdecls("w")}><w:t>1</w:t></w:r>',
                f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'):
        par._p.append(parse_xml(xml))


def _fill_hf(par, text):
    for part in HF_TOKEN.split(text):
        if part == "{page}":
            _fld(par, "PAGE")
        elif part == "{pages}":
            _fld(par, "NUMPAGES")
        elif part:
            par.add_run(part)


SHAPETYPE = (
    '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" '
    'path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/>'
    '<v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
    '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/>'
    '<v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/>'
    '<v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/>'
    '<v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" '
    'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
    '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" '
    'xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/>'
    '</v:shapetype>')


def _watermark_run(text):
    n = max(1, len(text))
    w = min(468.0, n * 80.0)
    h = w / (n * 0.8) if n * 80.0 > 468 else 100.0
    return parse_xml(
        f'<w:r {nsdecls("w")} xmlns:v="{V_NS}" xmlns:o="urn:schemas-microsoft-com:office:office" '
        f'xmlns:w10="urn:schemas-microsoft-com:office:word"><w:rPr><w:noProof/></w:rPr><w:pict>'
        f'{SHAPETYPE}<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" '
        f'type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;'
        f'width:{w:.0f}pt;height:{h:.0f}pt;rotation:315;z-index:-251658752;'
        f'mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
        f'mso-position-vertical:center;mso-position-vertical-relative:margin" '
        f'o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/>'
        f'<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" '
        f'string={quoteattr(text)}/><w10:wrap anchorx="margin" anchory="margin"/>'
        f'</v:shape></w:pict></w:r>')


def apply_page_setup(d, ps):
    """Write a PageSetup into the first section of a python-docx Document."""
    sec = d.sections[0]
    land = ps.orientation == "landscape"
    sec.orientation = WD_ORIENT.LANDSCAPE if land else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Mm(ps.eff_w), Mm(ps.eff_h)
    sec.top_margin, sec.bottom_margin = Mm(ps.top), Mm(ps.bottom)
    sec.left_margin, sec.right_margin = Mm(ps.left), Mm(ps.right)
    sp = sec._sectPr
    pgMar = sp.find(qn("w:pgMar"))

    if ps.line_numbers:
        pgMar.addnext(parse_xml(
            f'<w:lnNumType {nsdecls("w")} w:countBy="1" w:restart="newPage"/>'))
    if ps.border:
        sz = max(2, min(96, round(ps.border_width * 8)))
        color = "000000" if ps.border_color == "auto" else ps.border_color.lstrip("#")
        edges = "".join(
            f'<w:{e} w:val="{ps.border_style}" w:sz="{sz}" w:space="24" w:color="{color}"/>'
            for e in ("top", "left", "bottom", "right"))
        pgMar.addnext(parse_xml(f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">{edges}</w:pgBorders>'))
    _set_cols(sp, max(1, ps.columns), ps.col_gap)

    if ps.header or ps.watermark:
        hdr = sec.header
        hdr.is_linked_to_previous = False
        p = hdr.paragraphs[0]
        _fill_hf(p, ps.header)
        if ps.watermark:
            p._p.append(_watermark_run(ps.watermark))
        if ps.header:
            p.alignment = 1  # centre
    if ps.footer:
        ftr = sec.footer
        ftr.is_linked_to_previous = False
        p = ftr.paragraphs[0]
        _fill_hf(p, ps.footer)
        p.alignment = 1

    if ps.hyphenate:
        st = d.settings.element
        el = parse_xml(f'<w:autoHyphenation {nsdecls("w")}/>')
        anchor = st.find(qn("w:defaultTabStop"))
        if anchor is None:
            anchor = st.find(qn("w:zoom"))
        if anchor is not None:
            anchor.addnext(el)
        else:
            st.insert(0, el)


# --------------------------------------------------------------------------- read
def _hf_text(hf):
    """Header/footer text with PAGE / NUMPAGES fields turned back into tokens."""
    lines = []
    for p in hf.paragraphs:
        out, state = [], "text"
        for el in p._p.iter():
            tag = el.tag.split("}")[-1]
            if tag == "fldChar":
                t = el.get(qn("w:fldCharType"))
                state = {"begin": "instr", "separate": "result", "end": "text"}.get(t, state)
            elif tag == "instrText":
                code = (el.text or "").upper()
                out.append("{pages}" if "NUMPAGES" in code else "{page}" if "PAGE" in code else "")
            elif tag == "t" and state == "text":
                out.append(el.text or "")
        lines.append("".join(out))
    return "\n".join(x for x in lines if x).strip()


def _watermark_text(hdr):
    for el in hdr._element.iter():
        if el.tag == "{%s}textpath" % V_NS and el.get("string"):
            return el.get("string")
    return ""


def read_page_setup(d):
    """PageSetup from the first section of a Document (defaults for anything missing)."""
    ps = PageSetup()
    sec = d.sections[0]
    try:
        w, h = sec.page_width.mm, sec.page_height.mm
        ps.width_mm, ps.height_mm = min(w, h), max(w, h)
        ps.orientation = "landscape" if w > h else "portrait"
        ps.top, ps.bottom = sec.top_margin.mm, sec.bottom_margin.mm
        ps.left, ps.right = sec.left_margin.mm, sec.right_margin.mm
    except Exception:
        pass
    sp = sec._sectPr
    cols = sp.find(qn("w:cols"))
    if cols is not None:
        ps.columns = int(cols.get(qn("w:num"), 1) or 1)
        if cols.get(qn("w:space")):
            ps.col_gap = int(cols.get(qn("w:space"))) / 56.7
    if sp.find(qn("w:lnNumType")) is not None:
        ps.line_numbers = True
    pb = sp.find(qn("w:pgBorders"))
    if pb is not None and len(pb):
        ps.border = True
        top = pb[0]
        ps.border_style = top.get(qn("w:val"), "single")
        ps.border_width = int(top.get(qn("w:sz"), 8)) / 8
        col = top.get(qn("w:color"), "auto")
        ps.border_color = "auto" if col in ("auto", "000000") else col
    try:
        if not sec.header.is_linked_to_previous:
            ps.header = _hf_text(sec.header)
            ps.watermark = _watermark_text(sec.header)
        if not sec.footer.is_linked_to_previous:
            ps.footer = _hf_text(sec.footer)
    except Exception:
        pass
    ps.hyphenate = d.settings.element.find(qn("w:autoHyphenation")) is not None
    return ps
