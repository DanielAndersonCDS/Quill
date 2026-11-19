"""Write parsed Blocks into a QTextEdit cursor as ONE undoable edit."""
from PySide6.QtCore import Qt
from PySide6.QtGui import (QBrush, QColor, QTextBlockFormat, QTextCharFormat, QTextCursor,
                           QTextFrameFormat, QTextLength, QTextListFormat, QTextTableFormat)

# Custom char-format properties (must match quill.py's other *_PROP values)
MATH_PROP = 0x100000 + 3
BOX_PROP = 0x100000 + 4

CM = QTextCursor.MoveOperation


def run_format(r, size=None, force_bold=False, force_italic=False):
    f = QTextCharFormat()
    if size:
        f.setFontPointSize(size)
    if r.bold or force_bold:
        f.setFontWeight(700)
    if r.italic or force_italic:
        f.setFontItalic(True)
    if r.strike:
        f.setFontStrikeOut(True)
    if r.code:
        f.setFontFamilies(["Consolas", "Courier New", "monospace"])
    if r.sup:
        f.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
    elif r.sub:
        f.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSubScript)
    if r.math:
        f.setProperty(MATH_PROP, True)
    if r.box:
        f.setProperty(BOX_PROP, True)
    return f


def _insert_runs(cur, runs, **kw):
    for r in runs:
        cur.insertText(r.text, run_format(r, **kw))


def _make_empty_block(cur, new_bf):
    """Make sure the cursor sits in an empty block; return True if one was reused."""
    if not cur.block().text():
        return True
    at_start, at_end = cur.atBlockStart(), cur.atBlockEnd()
    cur.insertBlock(new_bf(), QTextCharFormat())
    if at_end:
        return False
    if not at_start:
        cur.insertBlock(new_bf(), QTextCharFormat())
    cur.movePosition(CM.PreviousBlock)
    return True


def make_table_format():
    fmt = QTextTableFormat()
    fmt.setBorder(1)
    fmt.setCellPadding(5)
    fmt.setCellSpacing(0)
    fmt.setBorderBrush(QBrush(QColor("#8c8fa1")))
    fmt.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
    fmt.setHeaderRowCount(1)
    try:
        fmt.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
    except Exception:
        pass
    return fmt


def _insert_table(cur, rows):
    ncols = max(len(r) for r in rows)
    table = cur.insertTable(len(rows), ncols, make_table_format())
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            runs = row[ci] if ci < len(row) else []
            c = table.cellAt(ri, ci).firstCursorPosition()
            _insert_runs(c, runs, force_bold=(ri == 0))
    end = table.lastCursorPosition()
    end.movePosition(CM.NextBlock)
    cur.setPosition(end.position())


def apply_blocks(cur, blocks, new_bf, heading_sizes):
    """Insert blocks at cur (a copy of the editor cursor); returns the final cursor.

    new_bf()        -> a fresh default QTextBlockFormat (supplied by quill.py)
    heading_sizes   -> {level: point size}
    """
    cur.beginEditBlock()
    try:
        if cur.hasSelection():
            cur.removeSelectedText()
        fresh = _make_empty_block(cur, new_bf)
        lists = {}
        for b in blocks:
            if b.kind == "table":
                if not fresh:
                    cur.insertBlock(new_bf(), QTextCharFormat())
                _insert_table(cur, b.rows)
                fresh, lists = True, {}
                continue

            bf = new_bf()
            kw = {}
            if b.kind == "heading":
                bf.setHeadingLevel(b.level)
                bf.setTopMargin(12)
                kw = dict(size=heading_sizes.get(b.level, 12), force_bold=True)
            elif b.kind == "math":
                bf.setAlignment(Qt.AlignmentFlag.AlignHCenter)
                bf.setTopMargin(4)
                bf.setBottomMargin(10)
            elif b.kind == "quote":
                bf.setLeftMargin(24)
                kw = dict(force_italic=True)
            elif b.kind == "code":
                bf.setLeftMargin(12)

            if fresh:
                cur.setBlockFormat(bf)
                cur.setBlockCharFormat(QTextCharFormat())
                fresh = False
            else:
                cur.insertBlock(bf, QTextCharFormat())

            if b.kind in ("bullet", "number"):
                key = b.level
                entry = lists.get(key)
                kind = b.kind
                if entry and entry[0] == kind:
                    entry[1].add(cur.block())
                else:
                    lf = QTextListFormat()
                    lf.setIndent(b.level + 1)
                    if kind == "number":
                        lf.setStyle(QTextListFormat.Style.ListDecimal)
                    else:
                        lf.setStyle(QTextListFormat.Style.ListDisc if b.level == 0
                                    else QTextListFormat.Style.ListCircle)
                    lists[key] = (kind, cur.createList(lf))
                for deeper in [k for k in lists if k > key]:
                    del lists[deeper]
            else:
                lists = {}

            _insert_runs(cur, b.runs, **kw)
    finally:
        cur.endEditBlock()
    return cur
