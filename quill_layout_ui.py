"""Qt side of page layout: dialogs, on-screen page decoration, PDF export."""
import math

from PySide6.QtCore import QLineF, QMarginsF, QRectF, QSizeF, Qt
from PySide6.QtGui import (QAbstractTextDocumentLayout, QColor, QFont, QFontMetricsF, QGuiApplication,
                           QPageLayout, QPageSize, QPainter, QPalette, QPdfWriter, QPen, QTextCursor,
                           QTextFormat)
from PySide6.QtWidgets import (QCheckBox, QColorDialog, QComboBox, QDialog, QDialogButtonBox,
                               QDoubleSpinBox, QFormLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QSpinBox, QVBoxLayout)

from quill_layout import BORDER_STYLES

PX_PER_MM = 96 / 25.4
BREAK = QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore


def tokens(text, page=1, pages=1):
    return text.replace("{pages}", str(pages)).replace("{page}", str(page))


# --------------------------------------------------------------------------- dialogs
class _Base(QDialog):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.form = QFormLayout()
        lay = QVBoxLayout(self)
        lay.addLayout(self.form)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                                        QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)


def _cm_spin(mm):
    s = QDoubleSpinBox()
    s.setRange(0.0, 20.0)
    s.setSingleStep(0.25)
    s.setDecimals(2)
    s.setSuffix(" cm")
    s.setValue(mm / 10)
    return s


def margins_dialog(parent, ps):
    d = _Base(parent, "Custom margins")
    spins = {k: _cm_spin(getattr(ps, k)) for k in ("top", "bottom", "left", "right")}
    for k, s in spins.items():
        d.form.addRow(k.capitalize() + ":", s)
    if d.exec():
        return ps.with_(**{k: s.value() * 10 for k, s in spins.items()})
    return None


def columns_dialog(parent, ps):
    d = _Base(parent, "Columns")
    n = QSpinBox()
    n.setRange(1, 6)
    n.setValue(ps.columns)
    gap = _cm_spin(ps.col_gap)
    d.form.addRow("Number of columns:", n)
    d.form.addRow("Spacing:", gap)
    d.form.addRow(QLabel("Columns show in the saved .docx and exported PDF;\n"
                         "the editing view stays single-column."))
    return ps.with_(columns=n.value(), col_gap=gap.value() * 10) if d.exec() else None


def header_footer_dialog(parent, ps):
    d = _Base(parent, "Header & footer")
    h, f = QLineEdit(ps.header), QLineEdit(ps.footer)
    d.form.addRow("Header:", h)
    d.form.addRow("Footer:", f)
    d.form.addRow(QLabel("Use {page} for the page number and {pages} for the page count,\n"
                         "e.g.  Page {page} of {pages}"))
    return ps.with_(header=h.text().strip(), footer=f.text().strip()) if d.exec() else None


def borders_dialog(parent, ps):
    d = _Base(parent, "Page borders")
    on = QCheckBox("Draw a border around every page")
    on.setChecked(ps.border)
    style = QComboBox()
    style.addItems(list(BORDER_STYLES))
    cur = next((k for k, v in BORDER_STYLES.items() if v == ps.border_style), "Single")
    style.setCurrentText(cur)
    width = QDoubleSpinBox()
    width.setRange(0.25, 6.0)
    width.setSingleStep(0.25)
    width.setSuffix(" pt")
    width.setValue(ps.border_width)
    color = {"v": ps.border_color}
    btn = QPushButton()

    def paint_btn():
        btn.setText("Automatic" if color["v"] == "auto" else "#" + color["v"].lstrip("#"))

    def pick():
        c = QColorDialog.getColor(QColor("#" + color["v"]) if color["v"] != "auto" else QColor("black"), d)
        if c.isValid():
            color["v"] = c.name()[1:].upper()
            paint_btn()

    btn.clicked.connect(pick)
    paint_btn()
    for label, w in (("", on), ("Style:", style), ("Width:", width), ("Colour:", btn)):
        d.form.addRow(label, w)
    if d.exec():
        return ps.with_(border=on.isChecked(), border_style=BORDER_STYLES[style.currentText()],
                        border_width=width.value(), border_color=color["v"])
    return None


# --------------------------------------------------------------------------- on-screen decoration
def _scan(ed):
    """Cache page-break positions and line-number starts per document revision/width."""
    ps, doc = ed.page, ed.document()
    key = (doc.revision(), ps.line_numbers, ed.viewport().width())
    if getattr(ed, "_scan_key", None) == key:
        return ed._scan_data
    nums, breaks, n = [], [], 1
    b = doc.begin()
    while b.isValid():
        if b.blockFormat().pageBreakPolicy() & BREAK:
            breaks.append(b.position())
        if ps.line_numbers and QTextCursor(b).currentTable() is None:
            nums.append((b.position(), n))
            n += max(1, b.lineCount())
        b = b.next()
    ed._scan_key, ed._scan_data = key, (nums, breaks)
    return ed._scan_data


def paint_decor(ed, p):
    """Draw watermark, border, header/footer, line numbers and page-break marks."""
    ps = ed.page
    vp = ed.viewport()
    w, h = vp.width(), vp.height()
    scroll = ed.verticalScrollBar().value()
    doc_h = ed.document().documentLayout().documentSize().height()
    ink = ed.palette().color(QPalette.ColorRole.Text)
    p.save()
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    ml, mr = ps.left * PX_PER_MM, ps.right * PX_PER_MM
    mt, mb = ps.top * PX_PER_MM, ps.bottom * PX_PER_MM

    if ps.watermark:
        f = QFont(ed.font())
        f.setBold(True)
        f.setPointSizeF(100)
        adv = max(1.0, QFontMetricsF(f).horizontalAdvance(ps.watermark))
        target = 0.75 * min(w, h) * 1.35
        f.setPointSizeF(max(18.0, min(150.0, 100.0 * target / adv)))
        c = QColor(ink)
        c.setAlpha(34)
        p.save()
        p.translate(w / 2, h / 2)
        p.rotate(-45)
        p.setFont(f)
        p.setPen(c)
        p.drawText(QRectF(-w, -h, 2 * w, 2 * h), Qt.AlignmentFlag.AlignCenter, ps.watermark)
        p.restore()

    if ps.border:
        c = ink if ps.border_color == "auto" else QColor("#" + ps.border_color.lstrip("#"))
        pen = QPen(c, max(1.0, ps.border_width * 96 / 72))
        pen.setStyle({"dashed": Qt.PenStyle.DashLine, "dotted": Qt.PenStyle.DotLine}.get(
            ps.border_style, Qt.PenStyle.SolidLine))
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        inset = 32
        p.drawRect(QRectF(inset, -scroll + inset, w - 2 * inset, doc_h - 2 * inset))

    small = QFont(ed.font())
    small.setPointSizeF(9)
    dim = QColor(ink)
    dim.setAlpha(150)
    p.setFont(small)
    p.setPen(dim)
    if ps.header:
        p.drawText(QRectF(ml, -scroll + mt * 0.3, w - ml - mr, 22),
                   Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop, tokens(ps.header))
    if ps.footer:
        p.drawText(QRectF(ml, doc_h - scroll - mb * 0.65, w - ml - mr, 22),
                   Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop, tokens(ps.footer))

    nums, breaks = _scan(ed)
    doc = ed.document()
    for pos, n in nums:
        b = doc.findBlock(pos)
        lay = b.layout()
        if lay is None:
            continue
        top = lay.position().y()
        if top - scroll > h or top - scroll < -400:
            continue
        for i in range(lay.lineCount()):
            ln = lay.lineAt(i)
            y = top + ln.y() - scroll
            if 0 <= y <= h:
                p.drawText(QRectF(0, y, max(8, ml - 10), ln.height()),
                           Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, str(n + i))
    pen = QPen(dim, 1, Qt.PenStyle.DashLine)
    for pos in breaks:
        b = doc.findBlock(pos)
        if b.layout() is None:
            continue
        y = b.layout().position().y() - scroll - 6
        if 0 <= y <= h:
            p.setPen(pen)
            p.drawLine(QLineF(ml, y, w - mr, y))
            p.setPen(dim)
            p.drawText(QRectF(ml, y - 16, w - ml - mr, 14), Qt.AlignmentFlag.AlignRight, "page break")
    p.restore()


# --------------------------------------------------------------------------- PDF
def export_pdf(qdoc, path, ps, font):
    """Paginate a clone of the document ourselves: margins, columns, header/footer,
    watermark and page border all come out in the PDF."""
    scr = QGuiApplication.primaryScreen()
    src = scr.logicalDotsPerInchX() if scr else 96.0
    mm = src / 25.4
    page_w, page_h = ps.eff_w * mm, ps.eff_h * mm
    ml, mr, mt, mb = ps.left * mm, ps.right * mm, ps.top * mm, ps.bottom * mm
    n = max(1, ps.columns)
    gap = ps.col_gap * mm
    col_w = max(50.0, (page_w - ml - mr - gap * (n - 1)) / n)
    body_h = max(50.0, page_h - mt - mb)

    doc = qdoc.clone()
    doc.setDocumentMargin(0)
    root = doc.rootFrame()
    ff = root.frameFormat()
    ff.setMargin(0)
    root.setFrameFormat(ff)
    doc.setPageSize(QSizeF(col_w, body_h))

    w = QPdfWriter(path)
    w.setPageSize(QPageSize(QSizeF(ps.eff_w, ps.eff_h), QPageSize.Unit.Millimeter))
    w.setPageMargins(QMarginsF(0, 0, 0, 0), QPageLayout.Unit.Millimeter)
    pages = max(1, doc.pageCount())
    sheets = math.ceil(pages / n)
    p = QPainter(w)
    sc = w.resolution() / src
    black = QColor("black")
    small = QFont(font)
    small.setPointSizeF(9)
    for sheet in range(sheets):
        if sheet:
            w.newPage()
        p.save()
        p.scale(sc, sc)
        if ps.watermark:
            f = QFont(font)
            f.setBold(True)
            f.setPixelSize(100)
            adv = max(1.0, QFontMetricsF(f).horizontalAdvance(ps.watermark))
            f.setPixelSize(int(max(30, min(260, 100 * 0.62 * page_w * 1.2 / adv))))
            c = QColor(160, 160, 160, 90)
            p.save()
            p.translate(page_w / 2, page_h / 2)
            p.rotate(-45)
            p.setFont(f)
            p.setPen(c)
            p.drawText(QRectF(-page_w, -page_h, 2 * page_w, 2 * page_h), Qt.AlignmentFlag.AlignCenter,
                       ps.watermark)
            p.restore()
        if ps.border:
            c = black if ps.border_color == "auto" else QColor("#" + ps.border_color.lstrip("#"))
            pen = QPen(c, ps.border_width * src / 72)
            pen.setStyle({"dashed": Qt.PenStyle.DashLine, "dotted": Qt.PenStyle.DotLine}.get(
                ps.border_style, Qt.PenStyle.SolidLine))
            p.setPen(pen)
            p.setBrush(Qt.BrushStyle.NoBrush)
            inset = 24 * src / 72
            p.drawRect(QRectF(inset, inset, page_w - 2 * inset, page_h - 2 * inset))
        p.setFont(small)
        p.setPen(QColor(90, 90, 90))
        if ps.header:
            p.drawText(QRectF(ml, mt * 0.35, page_w - ml - mr, 30),
                       Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                       tokens(ps.header, sheet + 1, sheets))
        if ps.footer:
            p.drawText(QRectF(ml, page_h - mb * 0.65, page_w - ml - mr, 30),
                       Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop,
                       tokens(ps.footer, sheet + 1, sheets))
        for c in range(n):
            idx = sheet * n + c
            if idx >= pages:
                break
            clip = QRectF(0, idx * body_h, col_w, body_h)
            p.save()
            p.translate(ml + c * (col_w + gap), mt - idx * body_h)
            p.setClipRect(clip)
            ctx = QAbstractTextDocumentLayout.PaintContext()
            ctx.clip = clip
            ctx.palette.setColor(QPalette.ColorRole.Text, black)
            doc.documentLayout().draw(p, ctx)
            p.restore()
        p.restore()
    p.end()
