# Excel Image Inserter

PyQt5 desktop utility for **batch-inserting images into Excel spreadsheets**.
Drop a folder of photos, pick anchor cells, set the crop preset and target
dimensions, and let the app pack everything into a single `.xlsx` with
correctly anchored thumbnails.

**Free & open-source (MIT).** If it saves you time on real work, please
consider [supporting development](./DONATE.md) — Telegram Stars, GitHub
Sponsors, Buy Me a Coffee, or crypto all welcome.

---

## Features

- Batch insert: many images → one Excel sheet in a single pass
- Crop presets: `1:1`, `4:3`, `3:2`, `16:9`, plus the vertical variants and a
  `None` pass-through
- One-cell anchor with sub-cell offset (image fits the cell, not the grid)
- Per-image grouping in the file list (collapse/expand)
- Optional re-sampling to control output `.xlsx` size — estimated and original
  file weights shown next to each entry
- Progress bar for long batches

## Stack

- Python 3.9+
- PyQt5 (GUI)
- openpyxl (Excel writer)
- Pillow (image probe + resize)

## Quick start

```bash
pip install -r requirements.txt
python main.py
```

## Build a standalone .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico main.py
```

The result lands in `dist/main.exe`.

## Roadmap

- Headless CLI mode (no GUI, drive from a JSON manifest)
- Drag-and-drop reorder of file list
- Save / reload "insert plan" for repeat batches

## License

MIT — see [LICENSE](./LICENSE).
