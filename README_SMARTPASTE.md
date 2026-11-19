# Quill: smart paste + themes

Unpack next to your existing quill.py (overwrite it). Layout:

    quill.py            (patched)
    quill_paste/        parser + LaTeX->Unicode (Qt-free) and apply.py (Qt)
    quill_themes.py     devkit JSON -> Quill palette
    themes/             nord.json, bloomberg.json  (drop more devkit themes here)
    tests/              python -m unittest discover -s tests -t .

## Using it
* Ctrl+V           auto-formats when the text looks like markdown / LaTeX / tables
* Ctrl+Shift+V     paste as plain text (no formatting)
* Ctrl+Alt+F       format the *selected* raw text already in the document
* Edit > Smart paste  toggles the behaviour (remembered)
* View > Theme     pick a theme; Ctrl+Shift+T cycles through all of them

## What gets converted
# headings, plain "1.1 Title" / short title lines, bullets, numbered lists, > quotes,
| pipe | tables |, tab-separated tables (real Word tables), **bold** *italic* `code`,
\[ display \] / \( inline \) / $$ $$ LaTeX incl. \boxed, aligned, \not, \mathbb, sub/superscripts,
markdown-escaped LaTeX (\\[ \\{ ... ) and &#x20; entities.
Math becomes Unicode text + italic/sub/sup runs, so it survives .docx save/load.
\boxed becomes a Word run border. Unknown LaTeX commands are left visible, never dropped.

---
# Layout tab, page features, icon, optional References panel

New files: quill_layout.py (page model + .docx round-trip), quill_layout_ui.py (dialogs, on-screen
preview, PDF export), assets/quill.ico + quill.png (app icon).

* Layout toolbar (View > Layout toolbar to hide): Text direction, Margins, Orientation, Size,
  Columns, Breaks, Line numbers, Hyphenation, Header & footer, Page number, Page borders,
  Watermark, plus Indent left/right and Spacing before/after for the selected paragraphs.
  The same commands are in the Layout and Insert menus.
* Page setup is saved into the .docx and read back when you open one.
* PDF export now uses your margins / size / orientation / columns / header / footer /
  watermark / border.
* References panel: close it with its x, Ctrl+Shift+R, the toolbar button, or View menu.
  Quill remembers whether it was open.

---
# Building Quill.exe (Windows)
Double-click build_exe.bat. It creates dist\Quill\Quill.exe with assets\quill.ico as the exe icon
(the same image is used for the window/taskbar). Share the whole dist\Quill folder (zip it).
Extra themes: drop .json files into a "themes" folder next to Quill.exe.
