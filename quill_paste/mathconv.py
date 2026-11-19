"""LaTeX -> styled Unicode runs. No Qt imports, so it can be unit-tested anywhere.

Word processors cannot render LaTeX, so we translate it into plain Unicode
characters plus a few style flags (italic variables, superscript, subscript,
boxed).  Those flags map 1:1 onto things a .docx run can store.
"""
import re
from dataclasses import dataclass, replace

LB = "\u2028"  # line break inside one paragraph


@dataclass(frozen=True)
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    code: bool = False
    sup: bool = False
    sub: bool = False
    math: bool = False
    box: bool = False
    pad: bool = False  # internal: spacing that may move behind a script


def _style(r):
    return (r.bold, r.italic, r.strike, r.code, r.sup, r.sub, r.math, r.box)


# --------------------------------------------------------------------------- tables
REL = {  # relations / arrows / logic: printed with spaces around them
    "in": "∈", "notin": "∉", "ni": "∋", "subseteq": "⊆", "subset": "⊂", "supseteq": "⊇",
    "supset": "⊃", "subsetneq": "⊊", "supsetneq": "⊋", "nsubseteq": "⊈", "sqsubseteq": "⊑",
    "leq": "≤", "le": "≤", "geq": "≥", "ge": "≥", "neq": "≠", "ne": "≠", "approx": "≈",
    "equiv": "≡", "sim": "∼", "cong": "≅", "propto": "∝", "mid": "∣", "nmid": "∤",
    "parallel": "∥", "perp": "⊥", "to": "→", "rightarrow": "→", "longrightarrow": "⟶",
    "Rightarrow": "⇒", "Longrightarrow": "⟹", "implies": "⇒", "Leftarrow": "⇐",
    "leftarrow": "←", "Leftrightarrow": "⇔", "leftrightarrow": "↔", "iff": "⇔",
    "impliedby": "⇐", "mapsto": "↦", "rightleftarrows": "⇄", "leftrightarrows": "⇆",
    "prec": "≺", "succ": "≻", "preceq": "⪯", "succeq": "⪰", "ll": "≪", "gg": "≫",
    "vdash": "⊢", "models": "⊨", "therefore": "∴", "because": "∵", "land": "∧",
    "wedge": "∧", "lor": "∨", "vee": "∨", "cup": "∪", "cap": "∩", "times": "×",
    "cdot": "·", "circ": "∘", "oplus": "⊕", "otimes": "⊗", "setminus": "∖", "pm": "±",
    "mp": "∓", "div": "÷", "bullet": "•", "star": "⋆", "ast": "∗", "uplus": "⊎",
    "sqcup": "⊔", "sqcap": "⊓", "triangle": "△", "bmod": "mod", "mod": "mod",
}
TIGHT = {  # printed without spacing
    "forall": "∀", "exists": "∃", "nexists": "∄", "neg": "¬", "lnot": "¬", "emptyset": "∅",
    "varnothing": "∅", "infty": "∞", "partial": "∂", "nabla": "∇", "ldots": "…", "dots": "…",
    "cdots": "⋯", "vdots": "⋮", "ddots": "⋱", "prime": "′", "angle": "∠", "top": "⊤",
    "bot": "⊥", "aleph": "ℵ", "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ", "wp": "℘",
    "square": "□", "blacksquare": "∎", "qed": "∎", "checkmark": "✓", "dagger": "†",
    "lceil": "⌈", "rceil": "⌉", "lfloor": "⌊", "rfloor": "⌋", "langle": "⟨", "rangle": "⟩",
    "lbrace": "{", "rbrace": "}", "lbrack": "[", "rbrack": "]", "lvert": "|", "rvert": "|",
    "vert": "|", "Vert": "‖", "sum": "∑", "prod": "∏", "bigcup": "⋃", "bigcap": "⋂",
    "int": "∫", "oint": "∮", "bigoplus": "⨁", "coprod": "∐", "bigvee": "⋁",
    "bigwedge": "⋀", "colon": ":", "degree": "°", "S": "§",
}
GREEK_LOW = dict(
    alpha="α", beta="β", gamma="γ", delta="δ", epsilon="ε", varepsilon="ε", zeta="ζ", eta="η",
    theta="θ", vartheta="ϑ", iota="ι", kappa="κ", lambda_="λ", mu="μ", nu="ν", xi="ξ", pi="π",
    varpi="ϖ", rho="ρ", varrho="ϱ", sigma="σ", varsigma="ς", tau="τ", upsilon="υ", phi="φ",
    varphi="ϕ", chi="χ", psi="ψ", omega="ω")
GREEK_LOW["lambda"] = "λ"
GREEK_UP = dict(Gamma="Γ", Delta="Δ", Theta="Θ", Lambda="Λ", Xi="Ξ", Pi="Π", Sigma="Σ",
                Upsilon="Υ", Phi="Φ", Psi="Ψ", Omega="Ω")
FUNCS = set("sin cos tan cot sec csc sinh cosh tanh arcsin arccos arctan log ln lg exp lim "
            "limsup liminf max min sup inf gcd lcm det dim deg ker arg Pr hom".split())
BB = {"Z": "ℤ", "N": "ℕ", "R": "ℝ", "Q": "ℚ", "C": "ℂ", "P": "ℙ", "H": "ℍ"}
NEG = {"in": "∉", "subseteq": "⊈", "subset": "⊄", "mid": "∤", "equiv": "≢", "=": "≠",
       "leq": "≰", "geq": "≱", "sim": "≁", "cong": "≇", "approx": "≉", "exists": "∄",
       "supseteq": "⊉", "ni": "∌", "parallel": "∦", "rightarrow": "↛", "Rightarrow": "⇏",
       "Leftrightarrow": "⇎", "leftrightarrow": "↮"}
ACCENT = {"overline": "\u0305", "bar": "\u0304", "hat": "\u0302", "widehat": "\u0302",
          "tilde": "\u0303", "widetilde": "\u0303", "vec": "\u20d7", "overrightarrow": "\u20d7",
          "dot": "\u0307", "ddot": "\u0308", "underline": "\u0332"}
SKIP_ARG = {"hspace", "vspace", "phantom", "label", "color", "tag", "hphantom", "vphantom"}
SKIP_NOARG = {"displaystyle", "textstyle", "scriptstyle", "limits", "nolimits", "nonumber",
              "notag", "quad_", "protect", "relax"}
SIZERS = {"left", "right", "bigl", "bigr", "Bigl", "Bigr", "big", "Big", "bigg", "Bigg",
          "biggl", "biggr", "middle"}
ALIGN_ENVS = {"aligned", "align", "align*", "alignat", "gather", "gathered", "gather*",
              "split", "equation", "equation*", "eqnarray", "multline", "multline*"}
MATRIX_WRAP = {"pmatrix": ("(", ")"), "bmatrix": ("[", "]"), "vmatrix": ("|", "|"),
               "Vmatrix": ("‖", "‖"), "Bmatrix": ("{", "}"), "matrix": ("", ""),
               "array": ("", ""), "smallmatrix": ("", ""), "cases": ("{", "")}
# characters after which +/- is a sign, not an operator
UNARY_PREV = set("([{,;:=<>≤≥≠∈∉⊆⊂⊇⇒⇔→⟶↔∧∨¬∀∃±∓+−×·∣∪∩∖∘⊕⊗")


def R(text, **kw):
    return Run(text, math=True, **kw)


class _Parser:
    def __init__(self, src, display=False, amp=""):
        self.s, self.i, self.n = src, 0, len(src)
        self.display, self.amp = display, amp
        self.script = 0
        self.upright = 0

    # ------------------------------------------------------------- low level
    def peek(self):
        return self.s[self.i] if self.i < self.n else ""

    def skip_ws(self):
        while self.i < self.n and self.s[self.i].isspace():
            self.i += 1

    def raw_group(self):
        """Raw text of {...} (balanced) or one character, after optional spaces."""
        self.skip_ws()
        if self.peek() != "{":
            c = self.peek()
            self.i += 1 if c else 0
            return c
        depth, j = 0, self.i
        while j < self.n:
            if self.s[j] == "{":
                depth += 1
            elif self.s[j] == "}":
                depth -= 1
                if depth == 0:
                    break
            j += 1
        out = self.s[self.i + 1:j]
        self.i = min(j + 1, self.n)
        return out

    def arg(self):
        """One math argument: a {group} or a single atom."""
        self.skip_ws()
        if self.peek() == "{":
            self.i += 1
            r = self.seq(True)
            if self.peek() == "}":
                self.i += 1
            return r
        return self.atom()

    def pad(self, sym):
        if self.script:
            return [R(sym)]
        return [R(" ", pad=True), R(sym), R(" ", pad=True)]

    @staticmethod
    def last_char(out):
        for r in reversed(out):
            if r.text.endswith(LB):
                return ""
            t = r.text.replace(" ", "")
            if t:
                return t[-1]
        return ""

    # ------------------------------------------------------------- sequences
    def seq(self, in_group=False):
        out = []
        while self.i < self.n:
            c = self.s[self.i]
            if c == "}":
                self.i += 1
                if in_group:
                    self.i -= 1
                    break
                continue
            start = self.i
            if c in "+-" and not self.script:
                self.i += 1
                last = self.last_char(out)
                g = "−" if c == "-" else "+"
                out += [R(g)] if (not last or last in UNARY_PREV) else self.pad(g)
            elif c == ",":
                self.i += 1
                thousands = re.match(r"\d{3}(?!\d)", self.s[self.i:self.i + 4]) is not None
                tight = self.script or (self.last_char(out).isdigit() and thousands)
                out += [R(",")] if tight else [R(","), R(" ", pad=True)]
            else:
                out += self.atom()
            self.scripts(out)
            if self.i == start:
                self.i += 1
        return out

    def scripts(self, out):
        while True:
            self.skip_ws()
            c = self.peek()
            if c not in ("^", "_"):
                return
            self.i += 1
            self.script += 1
            a = self.arg()
            self.script -= 1
            a = [replace(r, sup=(c == "^"), sub=(c == "_")) for r in a]
            tail = []
            while out and out[-1].pad:
                tail.insert(0, out.pop())
            out += a + tail

    def atom(self):
        c = self.peek()
        if not c:
            return []
        if c.isspace():
            self.i += 1
            return []
        if c in "^_":
            return []
        if c == "{":
            self.i += 1
            r = self.seq(True)
            if self.peek() == "}":
                self.i += 1
            return r
        if c == "\\":
            return self.command()
        self.i += 1
        return self.char(c)

    def char(self, c):
        if c == "$":
            return []
        if c.isalpha():
            return [R(c, italic=(c.isascii() and not self.upright))]
        if c in "=<>":
            return self.pad(c)
        if c == "*":
            return self.pad("∗")
        if c in ":;":
            return [R(c)] if self.script else [R(c), R(" ", pad=True)]
        if c == "-":
            return [R("−")]
        if c == "'":
            return [R("′")]
        if c == "~":
            return [R("\u00a0")]
        if c == "&":
            return [R(self.amp)] if self.amp else []
        return [R(c)]

    # ------------------------------------------------------------- commands
    def command(self):
        self.i += 1
        if self.i >= self.n:
            return [R("\\")]
        c = self.s[self.i]
        if not c.isalpha():
            self.i += 1
            return self.symbol_cmd(c)
        j = self.i
        while j < self.n and self.s[j].isalpha():
            j += 1
        name, self.i = self.s[self.i:j], j
        if name == "operatorname" and self.peek() == "*":
            self.i += 1
        return self.named(name)

    def symbol_cmd(self, c):
        if c == "\\":
            return [R(LB if self.display else " ")]
        if c == ",":
            return [R("\u2009")]
        if c in ";: ":
            return [R(" ")]
        if c in "!-":
            return []
        if c == "|":
            return [R("‖")]
        if c in "<>":
            return self.pad(c)
        return [R(c)]  # \{ \} \% \& \# \_ \$ ...

    def named(self, name):
        if name in REL:
            return self.pad(REL[name])
        if name in TIGHT:
            return [R(TIGHT[name])]
        if name in GREEK_LOW:
            return [R(GREEK_LOW[name], italic=not self.upright)]
        if name in GREEK_UP:
            return [R(GREEK_UP[name])]
        if name in FUNCS:
            return self.func(name)
        if name in SIZERS:
            self.skip_ws()
            if self.peek() == ".":
                self.i += 1
            return []
        if name in SKIP_ARG:
            self.raw_group()
            return []
        if name in SKIP_NOARG:
            return []
        if name in ("quad", "qquad"):
            return [R("\u2003" * (2 if name == "qquad" else 1))]
        h = getattr(self, "cmd_" + name, None)
        if h:
            return h()
        if name in ACCENT:
            return self.accent(ACCENT[name])
        return [R("\\" + name)]  # unknown: keep visible, never silently drop

    def func(self, name):
        out = [R(name)]
        j = self.i
        while j < self.n and self.s[j].isspace():
            j += 1
        if j < self.n and (self.s[j].isalnum() or self.s[j] == "\\") and self.s[j] not in "^_":
            out.append(R(" ", pad=True))
        return out

    def accent(self, comb):
        a = self.arg()
        return [replace(r, text="".join(ch + comb if ch != " " else ch for ch in r.text)) for r in a]

    def cmd_text(self):
        t = self.raw_group()
        t = re.sub(r"\\([%&#_${}])", r"\1", t).replace("\\ ", " ").replace("~", "\u00a0")
        t = re.sub(r"\\[a-zA-Z]+\s*", "", t).replace("$", "")
        return [R(t)] if t else []

    cmd_textrm = cmd_textsf = cmd_mbox = cmd_hbox = cmd_textnormal = cmd_text

    def cmd_textit(self):
        return [replace(r, italic=True) for r in self.cmd_text()]

    cmd_emph = cmd_textit

    def cmd_textbf(self):
        return [replace(r, bold=True) for r in self.cmd_text()]

    def upright_arg(self):
        self.upright += 1
        try:
            return self.arg()
        finally:
            self.upright -= 1

    def cmd_mathrm(self):
        return self.upright_arg()

    cmd_mathsf = cmd_mathtt = cmd_operatorname = cmd_mathop = cmd_mathrm

    def cmd_mathbf(self):
        return [replace(r, bold=True, italic=False) for r in self.upright_arg()]

    cmd_boldsymbol = cmd_bm = cmd_mathbf

    def cmd_mathit(self):
        return [replace(r, italic=True) for r in self.arg()]

    def cmd_mathcal(self):
        return self.arg()

    cmd_mathscr = cmd_mathfrak = cmd_mathcal

    def cmd_mathbb(self):
        t = self.raw_group()
        return [R(BB[ch]) if ch in BB else R(ch, bold=True) for ch in t if not ch.isspace()]

    def cmd_boxed(self):
        return [replace(r, box=True) for r in self.arg()]

    cmd_fbox = cmd_boxed

    @staticmethod
    def _wrap(runs):
        t = "".join(r.text for r in runs).strip()
        simple = len(t) <= 1 or re.fullmatch(r"[\w.]+", t)
        if simple:
            return runs
        return [R("(")] + runs + [R(")")]

    def cmd_frac(self):
        a, b = self.arg(), self.arg()
        return self._wrap(a) + [R("/")] + self._wrap(b)

    cmd_dfrac = cmd_tfrac = cmd_cfrac = cmd_frac

    def cmd_binom(self):
        a, b = self.arg(), self.arg()
        return [R("C(")] + a + [R(", ")] + b + [R(")")]

    cmd_dbinom = cmd_binom

    def cmd_sqrt(self):
        idx = []
        self.skip_ws()
        if self.peek() == "[":
            j = self.s.find("]", self.i)
            if j > 0:
                idx = [R(self.s[self.i + 1:j], sup=True)]
                self.i = j + 1
        return idx + [R("√")] + self._wrap(self.arg())

    def cmd_not(self):
        self.skip_ws()
        if self.peek() == "\\":
            save = self.i
            self.i += 1
            j = self.i
            while j < self.n and self.s[j].isalpha():
                j += 1
            nm = self.s[self.i:j]
            if nm in NEG:
                self.i = j
                return self.pad(NEG[nm])
            self.i = save
        elif self.peek() == "=":
            self.i += 1
            return self.pad("≠")
        a = self.atom()
        if a:
            a[-1] = replace(a[-1], text=a[-1].text + "\u0338")
        return a

    def cmd_textcolor(self):
        self.raw_group()
        return self.arg()

    def cmd_pmod(self):
        return [R(" (mod ")] + self.arg() + [R(")")]

    cmd_pod = cmd_pmod

    def cmd_begin(self):
        env = self.raw_group().strip()
        close = "\\end{" + env + "}"
        j = self.s.find(close, self.i)
        body = self.s[self.i:j] if j >= 0 else self.s[self.i:]
        self.i = j + len(close) if j >= 0 else self.n
        if env == "array":
            m = re.match(r"\s*\{[^}]*\}", body)
            body = body[m.end():] if m else body
        sub = _Parser(body, True, amp="" if env in ALIGN_ENVS else "\u2003\u2003")
        runs = sub.seq()
        if env in MATRIX_WRAP:
            lo, hi = MATRIX_WRAP[env]
            rows = split_lines(runs)
            out = []
            for k, row in enumerate(rows):
                pre = lo if (env != "cases" or k == 0) else " "
                out += ([R(LB)] if k else []) + ([R(pre + " ")] if pre.strip() else []) + row
                out += [R(" " + hi)] if hi else []
            return out
        return runs

    def cmd_end(self):
        self.raw_group()
        return []


def split_lines(runs):
    rows, cur = [], []
    for r in runs:
        if r.text == LB:
            rows.append(cur)
            cur = []
        else:
            cur.append(r)
    rows.append(cur)
    return rows


def tidy(runs):
    """Merge neighbouring runs of equal style and normalise spacing."""
    runs = [replace(r, pad=False) for r in runs if r.text]
    for _ in range(2):
        out = []
        for r in runs:
            if out and out[-1].text.endswith(" ") and r.text.startswith(" "):
                r = replace(r, text=r.text.lstrip(" "))
            if out and r.text == LB:
                out[-1] = replace(out[-1], text=out[-1].text.rstrip(" "))
            if out and out[-1].text == LB:
                r = replace(r, text=r.text.lstrip(" "))
            if not r.text:
                continue
            if out and _style(out[-1]) == _style(r):
                out[-1] = replace(out[-1], text=out[-1].text + r.text)
            else:
                out.append(r)
        runs = [r for r in out if r.text]
    if runs:
        runs[0] = replace(runs[0], text=runs[0].text.lstrip(" "))
        runs[-1] = replace(runs[-1], text=runs[-1].text.rstrip(" "))
    return [replace(r, text=re.sub(r" {2,}", " ", r.text)) for r in runs if r.text]


def convert(src, display=False):
    """LaTeX source -> list[Run]; LB runs mark line breaks."""
    return tidy(_Parser(src.strip(), display).seq())


def plain(runs):
    return "".join(r.text for r in runs)
