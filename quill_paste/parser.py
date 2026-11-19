"""Turn pasted text (markdown-ish, LaTeX, tab/pipe tables, plain notes) into a
neutral list of Blocks.  No Qt here; apply.py turns Blocks into document edits.
"""
import re
from dataclasses import dataclass, field, replace

from .mathconv import LB, Run, convert, tidy

MAX_MATH_PH = 0x1000


@dataclass
class Block:
    kind: str                       # heading para bullet number quote math code table
    runs: list = field(default_factory=list)
    level: int = 0                  # heading level (1-4) or list depth (0-2)
    rows: list = field(default_factory=list)  # table: rows -> cells -> runs


# --------------------------------------------------------------------------- cleanup
def normalize(text):
    t = text.replace("\r\n", "\n").replace("\r", "\n").replace("\u200b", "").replace("\ufeff", "")
    t = re.sub(r"&#x([0-9a-fA-F]+);", lambda m: chr(int(m.group(1), 16)), t)
    t = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), t)
    t = t.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
    if re.search(r"\\\\[\[(]", t):  # markdown-escaped LaTeX:  \\[ x \\{ ... -> \[ x \{
        t = re.sub(r"\\\\(?=[\[\](){}<>_&%#$|]|begin\b|end\b)", r"\\", t)
    return "\n".join(ln.rstrip() for ln in t.split("\n"))


INLINE_MATH = re.compile(
    r"\\\((?P<a>.+?)\\\)|\\\[(?P<b>.+?)\\\]|\$\$(?P<c>.+?)\$\$|(?<![\\$\w])\$(?P<d>[^$\n]+?)\$(?!\w)")


def _plausible_dollar(s):
    if re.search(r"[\\^_{}]", s):
        return True
    return re.fullmatch(r"[A-Za-z](_?\w)?|[A-Za-z]\s*[=<>]\s*\w+", s.strip()) is not None


# --------------------------------------------------------------------------- inline
EMPH = [
    (re.compile(r"\*\*\*(?=\S)(.+?)(?<=\S)\*\*\*"), (True, True, False)),
    (re.compile(r"\*\*(?=[^\s*])(.+?)\*\*(?!\*)"), (True, False, False)),
    (re.compile(r"(?<!\w)__(?=\S)(.+?)(?<=\S)__(?!\w)"), (True, False, False)),
    (re.compile(r"~~(?=\S)(.+?)(?<=\S)~~"), (False, False, True)),
    (re.compile(r"\*(?=[^\s*])(.+?)(?<=[^\s*])\*"), (False, True, False)),
    (re.compile(r"(?<!\w)_(?=[^\s_])(.+?)(?<=[^\s_])_(?!\w)"), (False, True, False)),
]


def inline(text, bold=False, italic=False, display_math=False):
    """Inline markdown + LaTeX -> runs."""
    store = []

    def keep(runs):
        store.append(runs)
        return chr(0xE000 + len(store) - 1)

    text = re.sub(r"`([^`\n]+)`", lambda m: keep([Run(m.group(1), code=True)]), text)

    def math(m):
        src = next(g for g in m.groups() if g is not None)
        if m.group("d") is not None and not _plausible_dollar(src):
            return m.group(0)
        if len(store) >= MAX_MATH_PH:
            return m.group(0)
        return keep(convert(src, display_math))

    text = INLINE_MATH.sub(math, text)
    text = re.sub(r"\\([\\`*_{}\[\]()#+\-.!<>|~$%&])", lambda m: chr(0xF000 + ord(m.group(1))), text)
    text = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)",
                  lambda m: m.group(1) if m.group(1) == m.group(2) else f"{m.group(1)} ({m.group(2)})", text)
    text = re.sub(r"<br\s*/?>", LB, text, flags=re.I)
    return tidy_text(_emph(text, (bold, italic, False), store))


def _emph(s, fl, store):
    best = None
    for rx, add in EMPH:
        m = rx.search(s)
        if m and (best is None or m.start() < best[0].start()):
            best = (m, add)
    if best is None:
        return _leaf(s, fl, store)
    m, add = best
    inner = (fl[0] or add[0], fl[1] or add[1], fl[2] or add[2])
    return _leaf(s[:m.start()], fl, store) + _emph(m.group(1), inner, store) + \
        _emph(s[m.end():], fl, store)


def _leaf(s, fl, store):
    out, buf = [], []

    def flush():
        if buf:
            txt = "".join(buf).replace("**", "")
            if txt:
                out.append(Run(txt, bold=fl[0], italic=fl[1], strike=fl[2]))
            buf.clear()

    for ch in s:
        o = ord(ch)
        if 0xE000 <= o < 0xF000:
            flush()
            for r in store[o - 0xE000]:
                out.append(replace(r, bold=r.bold or fl[0], italic=r.italic or fl[1],
                                   strike=r.strike or fl[2]))
        elif 0xF000 <= o < 0xF100:
            buf.append(chr(o - 0xF000))
        else:
            buf.append(ch)
    flush()
    return out


def tidy_text(runs):
    out = []
    for r in runs:
        if not r.text:
            continue
        if out and out[-1] == replace(r, text=out[-1].text):
            out[-1] = replace(out[-1], text=out[-1].text + r.text)
        else:
            out.append(r)
    return out


# --------------------------------------------------------------------------- block regexes
HEAD_MD = re.compile(r"^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$")
BULLET = re.compile(r"^(\s*)[-*+•–]\s+(\S.*)$")
NUMBERED = re.compile(r"^(\s*)(\d{1,3})[.)]\s+(\S.*)$")
DOTTED = re.compile(r"^(\d{1,3}(?:\.\d{1,3})+)\.?\s+(\S.*)$")
QUOTE = re.compile(r"^\s*>\s?(.*)$")
HR = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
PIPE_SEP = re.compile(r"^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$")
STOP = {"yes", "no", "therefore", "hence", "thus", "similarly", "notice", "remember", "note",
        "and", "but", "so", "or", "not", "then", "next", "first", "second", "finally",
        "given", "where", "here", "otherwise", "true", "false"}


def _mask(line):
    store = []

    def keep(m):
        store.append(m.group(0))
        return chr(0xE000 + len(store) - 1)

    return INLINE_MATH.sub(keep, line), store


def _unmask(cell, store):
    return "".join(store[ord(c) - 0xE000] if 0xE000 <= ord(c) < 0xF000 else c for c in cell)


def split_pipe(line):
    masked, store = _mask(line.strip())
    if masked.startswith("|"):
        masked = masked[1:]
    if masked.endswith("|") and not masked.endswith("\\|"):
        masked = masked[:-1]
    return [_unmask(c, store).strip().replace("\\|", "|") for c in re.split(r"(?<!\\)\|", masked)]


def _is_pipe_row(line):
    return line.strip().startswith("|") and line.count("|") >= 2


def _tab_cells(line):
    s = line.strip("\t ")
    return [c.strip() for c in s.split("\t")] if "\t" in s else None


def _heading_shape(line):
    """Short, title-like, unpunctuated line (used for plain-text headings)."""
    s = line.strip()
    words = s.split()
    if not (1 <= len(words) <= 9) or not (4 <= len(s) <= 70) or len(s) < 6:
        return False
    if not (s[0].isupper() or s[0].isdigit()) or re.search(r"[.,;:!?…]$", s):
        return False
    if re.search(r"[\\|\t=→⇒]|[\uE000-\uF8FF]", s) or s.isdigit() or words[0].lower() in STOP:
        return False
    return not (BULLET.match(line) or NUMBERED.match(line) or QUOTE.match(line))


# --------------------------------------------------------------------------- detection
def looks_structured(text):
    """True when pasted text has enough markup that auto-formatting is welcome."""
    t = normalize(text)
    score = 0
    if re.search(r"\\[\[(]|\$\$|\\begin\{|\\(?:frac|forall|exists|mid|subseteq|Rightarrow|boxed)\b", t):
        score += 3
    if re.search(r"(?m)^\s{0,3}#{1,6}\s+\S", t):
        score += 2
    lines = t.split("\n")
    if any(_is_pipe_row(a) and PIPE_SEP.match(b) for a, b in zip(lines, lines[1:])):
        score += 3
    if any("\t" in a.strip() and "\t" in b.strip() for a, b in zip(lines, lines[1:])):
        score += 2
    if re.search(r"\*\*[^*\n]+\*\*", t):
        score += 1
    if len(re.findall(r"(?m)^\s*[-*+•]\s+\S", t)) >= 2:
        score += 2
    if len(re.findall(r"(?m)^\s*\d{1,3}[.)]\s+\S", t)) >= 2:
        score += 1
    if re.search(r"(?m)^\d{1,3}(\.\d{1,3})+\s+[A-Z]", t):
        score += 1
    return score >= 2


# --------------------------------------------------------------------------- blocks
class _Blocks:
    def __init__(self, text):
        self.lines = normalize(text).split("\n")
        self.n = len(self.lines)
        self.out = []
        self.stack = []  # list indent stack

    def nonblank_after(self, i):
        j = i + 1
        while j < self.n and not self.lines[j].strip():
            j += 1
        return j

    def add(self, kind, runs, level=0):
        if kind not in ("bullet", "number"):
            self.stack = []
        if runs or kind == "math":
            self.out.append(Block(kind, runs, level))

    # --- display math
    def display_math(self, i):
        s = self.lines[i].strip()
        pairs = (("\\[", "\\]"), ("$$", "$$"))
        for op, cl in pairs:
            if not s.startswith(op):
                continue
            rest = s[len(op):]
            k = rest.find(cl)
            if k >= 0:
                return rest[:k], i + 1, rest[k + len(cl):].strip()
            buf = [rest]
            for j in range(i + 1, self.n):
                k = self.lines[j].find(cl)
                if k >= 0:
                    buf.append(self.lines[j][:k])
                    return "\n".join(buf), j + 1, self.lines[j][k + len(cl):].strip()
                buf.append(self.lines[j])
            return None
        m = re.match(r"\\begin\{(aligned|align\*?|gather\*?|equation\*?|split|eqnarray\*?)\}", s)
        if m:
            end = "\\end{" + m.group(1) + "}"
            buf = []
            for j in range(i, self.n):
                buf.append(self.lines[j])
                if end in self.lines[j]:
                    return "\n".join(buf), j + 1, ""
        m = re.fullmatch(r"\\\(((?:(?!\\\)).)+)\\\)", s)
        if m:
            return m.group(1), i + 1, ""
        return None

    # --- tables
    def pipe_table(self, i):
        if not (_is_pipe_row(self.lines[i]) and i + 1 < self.n and PIPE_SEP.match(self.lines[i + 1])):
            return None
        rows = [split_pipe(self.lines[i])]
        j = i + 2
        while j < self.n and _is_pipe_row(self.lines[j]):
            rows.append(split_pipe(self.lines[j]))
            j += 1
        return rows, j

    def tab_table(self, i):
        first = _tab_cells(self.lines[i])
        if not first or sum(1 for c in first if c) < 2:
            return None
        rows, j = [first], i + 1
        while j < self.n:
            c = _tab_cells(self.lines[j])
            if not c or len(c) != len(first):
                break
            rows.append(c)
            j += 1
        return (rows, j) if len(rows) >= 2 else None

    def add_table(self, rows):
        width = max(len(r) for r in rows)
        grid = []
        for k, r in enumerate(rows):
            cells = [inline(c, bold=(k == 0)) for c in r] + [[] for _ in range(width - len(r))]
            grid.append(cells)
        self.stack = []
        self.out.append(Block("table", rows=grid))

    # --- headings that are not marked with #
    def numbered_heading(self, i, m):
        text = m.group(3).strip()
        if re.search(r"[.,;:!?]$", text) or len(text) > 90 or not text[0].isupper():
            return False
        if INLINE_MATH.search(text) or m.group(1):
            return False

        def numbered(ln):
            return ln is not None and NUMBERED.match(ln) is not None

        prev = i - 1
        while prev >= 0 and not self.lines[prev].strip():
            prev -= 1
        nxt = self.nonblank_after(i)
        return not (prev >= 0 and numbered(self.lines[prev])) and \
            not (nxt < self.n and numbered(self.lines[nxt]))

    def run(self):
        i, first_seen = 0, False
        while i < self.n:
            raw = self.lines[i]
            s = raw.strip()
            if not s:
                i += 1
                continue
            at_start = not first_seen
            first_seen = True

            if s.startswith("```"):
                buf, j = [], i + 1
                while j < self.n and not self.lines[j].strip().startswith("```"):
                    buf.append(self.lines[j])
                    j += 1
                self.add("code", [Run(LB.join(buf), code=True)])
                i = j + 1
                continue

            dm = self.display_math(i)
            if dm:
                src, i, tail = dm
                self.add("math", convert(src, True))
                if tail:
                    self.add("para", inline(tail))
                continue

            pt = self.pipe_table(i) or self.tab_table(i)
            if pt:
                self.add_table(pt[0])
                i = pt[1]
                continue

            if HR.match(s):
                i += 1
                continue
            m = HEAD_MD.match(raw)
            if m:
                self.add("heading", inline(m.group(2)), min(len(m.group(1)), 4))
                self.out[-1].runs = [replace(r, bold=False) for r in self.out[-1].runs]
                i += 1
                continue
            m = QUOTE.match(raw)
            if m:
                self.add("quote", inline(m.group(1)))
                i += 1
                continue
            m = DOTTED.match(s)
            if m and s[len(m.group(1)):].strip()[:1].isupper() and len(s) <= 120 and \
                    not INLINE_MATH.search(s):
                self.heading_from(m.group(2), 2 + min(m.group(1).count("."), 2))
                i += 1
                continue
            m = NUMBERED.match(raw)
            if m and self.numbered_heading(i, m):
                self.heading_from(m.group(3), 2)
                i += 1
                continue
            m = BULLET.match(raw) or NUMBERED.match(raw)
            if m and not s.startswith("**"):
                kind = "number" if NUMBERED.match(raw) else "bullet"
                indent = len(m.group(1).expandtabs(4))
                while self.stack and indent < self.stack[-1]:
                    self.stack.pop()
                if not self.stack or indent > self.stack[-1]:
                    self.stack.append(indent)
                self.add(kind, inline(m.group(m.lastindex)), min(len(self.stack) - 1, 2))
                i += 1
                continue
            if _heading_shape(raw) and not s.startswith("**"):
                prev_is_head = i > 0 and _heading_shape(self.lines[i - 1])
                nxt = self.nonblank_after(i)
                if (at_start or not prev_is_head) and nxt < self.n:
                    self.heading_from(s, 1 if at_start else 3)
                    i += 1
                    continue
            self.add("para", inline(s))
            i += 1
        return self.out

    def heading_from(self, text, level):
        runs = [replace(r, bold=False) for r in inline(text)]
        self.add("heading", runs, level)


def parse(text):
    """Pasted text -> list[Block]."""
    return _Blocks(text).run()
