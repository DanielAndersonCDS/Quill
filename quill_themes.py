"""Load devkit-style theme JSON files and expand them into Quill's full palette.

A theme file only needs the devkit tokens:
    name bg surface surface_alt border text text_dim accent accent_text
    success warning danger info            (radius / font are ignored)
Drop any such .json into the `themes/` folder next to quill.py and it shows up
under View > Theme.  Filename (without .json) becomes the theme id.
"""
import json
from pathlib import Path

REQUIRED = ("bg", "surface", "surface_alt", "border", "text", "text_dim", "accent",
            "success", "warning", "danger", "info")


def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _hex(c):
    return "#%02x%02x%02x" % tuple(max(0, min(255, round(v))) for v in c)


def mix(a, b, t):
    """Blend colour a toward b by t (0..1)."""
    ca, cb = _rgb(a), _rgb(b)
    return _hex(tuple(x + (y - x) * t for x, y in zip(ca, cb)))


def luminance(h):
    r, g, b = _rgb(h)
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255


def palette_from_tokens(tokens, theme_id):
    """Expand devkit tokens into every colour key Quill's QSS and highlighter use."""
    missing = [k for k in REQUIRED if k not in tokens]
    if missing:
        raise ValueError("missing tokens: " + ", ".join(missing))
    t = tokens
    dark = luminance(t["bg"]) < 0.5
    black_white = "#000000" if dark else "#ffffff"
    r, g, b = _rgb(t["accent"])
    return dict(
        name=theme_id, label=t.get("name", theme_id), dark=dark,
        base=t["bg"], mantle=t["surface"], surface0=t["surface_alt"], surface1=t["border"],
        surface2=mix(t["border"], t["text"], 0.2),
        crust=mix(t["surface"], "#000000", 0.35 if dark else 0.06),
        text=t["text"], subtext1=mix(t["text"], t["text_dim"], 0.5), subtext0=t["text_dim"],
        overlay0=mix(t["text_dim"], t["border"], 0.5),
        mauve=t["accent"], lavender=mix(t["accent"], t["text"], 0.35),
        accent_text=t.get("accent_text", t["bg"]),
        accent_bg=f"rgba({r},{g},{b},0.18)",
        red=t["danger"], green=t["success"], yellow=t["warning"], blue=t["info"],
        peach=mix(t["warning"], t["danger"], 0.5), teal=mix(t["success"], t["info"], 0.5),
        sapphire=mix(t["info"], black_white, 0.15), sky=mix(t["info"], t["text"], 0.3),
        pink=mix(t["accent"], t["danger"], 0.5), maroon=mix(t["danger"], t["text"], 0.25),
        flamingo=mix(t["danger"], t["text"], 0.45), rosewater=mix(t["danger"], t["text"], 0.6),
    )


def load_themes(folder, skip=("mocha", "latte")):
    """Return ({id: palette}, [problems]) for every valid *.json in folder."""
    found, problems = {}, []
    folder = Path(folder)
    if not folder.is_dir():
        return found, problems
    for f in sorted(folder.glob("*.json")):
        if f.stem.lower() in skip:
            continue  # Mocha/Latte are built in with a hand-tuned palette
        try:
            found[f.stem] = palette_from_tokens(json.loads(f.read_text(encoding="utf-8")), f.stem)
        except (OSError, ValueError, KeyError) as e:
            problems.append(f"{f.name}: {e}")
    return found, problems
