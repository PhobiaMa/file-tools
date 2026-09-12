# File Tools — desktop utility

A small Windows/Mac/Linux desktop app (Python + tkinter, no heavy framework) with three everyday tools:

1. **Bulk rename** — find/replace, prefix/suffix, numbering (`photo_001`, `photo_002`…), extension filter, **live preview** before applying, collision-safe two-pass rename.
2. **Merge PDFs** — add files, reorder with ↑/↓, merge into one PDF.
3. **CSV → Excel** — converts one or many CSV files; auto-detects `;` `,` tab `|` separators, turns numbers into real numbers, bold header, filters, auto column widths.

## Run it
```bash
pip install -r requirements.txt
python file_tools.py
```

Can be packaged into a single `.exe` (PyInstaller) so the client needs nothing installed.

## Typical requests this covers
Batch file conversions, folder clean-up scripts, PDF split/merge/watermark, Excel report generators, small internal tools with a simple GUI.

---

*FR — Petit outil bureau : renommage en masse avec aperçu, fusion de PDF, conversion CSV → Excel formaté. Livrable en `.exe` sans installation.*
