# Excel Image Inserter

**Batch-insert and arrange images into Excel sheets — no macros, no plugins, just a desktop app.**

[![Build](https://github.com/moskva1988/excel-image-inserter/actions/workflows/build.yml/badge.svg)](https://github.com/moskva1988/excel-image-inserter/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey.svg)](https://github.com/moskva1988/excel-image-inserter/releases)

<!-- Screenshots will be added after UI modernization (Phase 6) -->
![Screenshot](docs/screenshots/main.png)

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Building from Source](#building-from-source)
- [Project Structure](#project-structure)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [Support the Project](#support-the-project)
- [License](#license)
- [Author](#author)

---

## Overview

Excel Image Inserter is a cross-platform desktop utility that lets you load a collection of images, organize them into named groups, configure crop ratios and target dimensions, and export everything into an `.xlsx` file with a single click.

Images are placed in a configurable grid — you choose the number of columns, the starting cell, and whether images float over existing cell content or are anchored inside cells. Each sheet gets a collapsible table of contents at the top with hyperlinks that jump directly to each image group, making large files easy to navigate.

The tool is aimed at anyone who regularly needs to produce image catalogs, product sheets, inspection reports, or photo archives in Excel format, and is tired of doing it manually cell by cell.

---

## Features

### Image management
- Batch-add images via drag-and-drop or a file picker
- Organize images into named **groups (categories)**
  - Add, remove, and rename groups
  - Drag images between groups
- Three list view modes: **List**, **Details**, **Stack**
- Stats bar showing image count, original size, estimated compressed size, and resolution

### Image processing
- **Crop presets**: 1:1, 4:3, 3:2, 16:9, 3:4, 2:3, 9:16, or no crop
- **Scale to pixels**: set a maximum width and/or height; aspect ratio is preserved
- All processing happens at export time — source files are never modified

### Excel output
- Place images in a **grid**: configurable column count and starting cell
- **Overlay mode**: images can float over existing cell data
- **Collapsible TOC** inserted at the top of each sheet with hyperlinks to every group
- Export to a **new `.xlsx` file** or **append to an existing file**
- **Create a new sheet** and control where it is inserted (by name, after a chosen sheet)
- ⚠ `.xlsx`-only mode (Excel 2007+ format); older `.xls` files are not supported

### Application
- Cross-platform: Windows and macOS
- Help **"?"** button with an About dialog (version, build number)
- No internet connection required; no telemetry

---

## Installation

### For users — download a pre-built binary

No Python installation required.

1. Go to the [Releases page](https://github.com/moskva1988/excel-image-inserter/releases).
2. Download the file for your OS:
   - **Windows**: `ExcelImageInserter.exe` — double-click to run
   - **macOS**: `Excel-Image-Inserter-macOS.zip` — unzip and move `Excel Image Inserter.app` to your Applications folder
3. On macOS, if Gatekeeper blocks the app: right-click → Open → Open.

### For developers — run from source

**Prerequisites:** Python 3.11 or newer.

```bash
# Clone the repository
git clone https://github.com/moskva1988/excel-image-inserter.git
cd excel-image-inserter

# Create and activate a virtual environment
python3 -m venv venv
source venv/bin/activate        # macOS / Linux
# venv\Scripts\activate          # Windows

# Install dependencies
pip install -r requirements.txt

# Launch the app
python main.py
```

**Dependencies** (`requirements.txt`):

| Package | Version | Purpose |
|---|---|---|
| PyQt5 | ≥ 5.15 | GUI framework |
| openpyxl | ≥ 3.1 | Excel file read/write |
| Pillow | ≥ 10.0 | Image processing |

---

## Usage

<!-- Screenshots will be added after UI modernization (Phase 6) -->

### Step 1 — Launch the app

Open `ExcelImageInserter.exe` (Windows) or `Excel Image Inserter.app` (macOS), or run `python main.py` from the project directory.

<!-- ![Step 1](docs/screenshots/step-1.png) -->

### Step 2 — Open or create an Excel file

Click **Browse** next to the "Excel File" field to select an existing `.xlsx` file, or leave the field empty to create a new file at export time.

<!-- ![Step 2](docs/screenshots/step-2.png) -->

### Step 3 — Add images

Drag image files onto the image list, or use the **Add Images** button to open a file picker. Supported formats are those handled by Pillow (JPEG, PNG, BMP, TIFF, WebP, and others).

<!-- ![Step 3](docs/screenshots/step-3.png) -->

### Step 4 — Organize into groups

Use the **Add Group** button to create a named category. Select images in the list and drag them into the target group. Groups can be renamed (double-click) or removed. The order of groups determines the order they appear in the Excel sheet and TOC.

<!-- ![Step 4](docs/screenshots/step-4.png) -->

### Step 5 — Configure image processing

- **Crop ratio**: choose a preset (1:1, 4:3, 16:9, etc.) or leave as "None" to keep the original aspect ratio.
- **Scale**: set a maximum width and/or height in pixels. Images larger than the limit are scaled down; smaller images are not upscaled.

<!-- ![Step 5](docs/screenshots/step-5.png) -->

### Step 6 — Configure grid and sheet placement

- **Columns**: how many images per row.
- **Starting cell**: the top-left cell where the first image is placed (e.g. `A2`).
- **Overlay on cells**: when checked, images float over existing cell content instead of being anchored inside cells.
- **Sheet**: choose an existing sheet from the dropdown, or type a new sheet name. For new sheets, select the sheet it should be inserted after.

<!-- ![Step 6](docs/screenshots/step-6.png) -->

### Step 7 — Export

Click **Export to Excel**. A progress bar tracks the operation. When complete, the app reports the output file path.

<!-- ![Step 7](docs/screenshots/step-7.png) -->

---

## Building from Source

Install [PyInstaller](https://pyinstaller.org/) in addition to the project dependencies:

```bash
pip install pyinstaller
```

**Windows** — produces a single portable `.exe`:

```bash
pyinstaller --onefile --windowed --name "ExcelImageInserter" --icon icon.ico main.py
# Output: dist/ExcelImageInserter.exe
```

**macOS** — produces a `.app` bundle:

```bash
# First generate the .icns icon (requires macOS iconutil)
mkdir -p icon.iconset
python3 -c "
from PIL import Image
img = Image.open('icon.png')
sizes = {
    'icon_16x16.png': 16, 'icon_16x16@2x.png': 32,
    'icon_32x32.png': 32, 'icon_32x32@2x.png': 64,
    'icon_128x128.png': 128, 'icon_128x128@2x.png': 256,
    'icon_256x256.png': 256, 'icon_256x256@2x.png': 512,
    'icon_512x512.png': 512, 'icon_512x512@2x.png': 1024,
}
for name, size in sizes.items():
    img.resize((size, size), Image.LANCZOS).save(f'icon.iconset/{name}')
"
iconutil -c icns icon.iconset

# Build the .app
pyinstaller --windowed --name "Excel Image Inserter" --icon icon.icns main.py
# Output: dist/Excel Image Inserter.app
```

**GitHub Actions** handles both builds automatically on every tag push matching `v*`. See [`.github/workflows/build.yml`](.github/workflows/build.yml) for the full workflow. Release artifacts (`ExcelImageInserter.exe` and `Excel-Image-Inserter-macOS.zip`) are attached to the GitHub Release automatically.

---

## Project Structure

```
excel-image-inserter/
├── main.py                  # Application entry point (monolithic, ~single file)
├── icon.ico                 # Windows application icon
├── icon.png                 # Source icon (used to generate .icns for macOS)
├── requirements.txt         # Python dependencies
├── LICENSE                  # MIT License
├── .github/
│   └── workflows/
│       └── build.yml        # CI/CD: build Windows .exe and macOS .app on tag push
└── docs/
    └── screenshots/         # Placeholder — screenshots added in Phase 6
```

> Note: the codebase is currently a single `main.py` file. A modular split into separate files (UI, worker, export logic, etc.) is planned — see [Roadmap](#roadmap).

---

## Roadmap

- **Modular code split** — extract UI panels, export worker, and image-processing logic into separate modules (in progress)
- **Modern Fluent UI** with Light and Dark theme support
- **Image Batch Processor module** — folder-to-folder workflow: resize, crop, metadata editing, date/image watermarks — independent of Excel output
- **Screenshots** added to documentation after the UI refresh (Phase 6)

---

## Contributing

Issues and pull requests are welcome. Before submitting a larger change, please open an issue to discuss the approach.

- Follow the existing code style (PEP 8, PyQt5 patterns already in `main.py`)
- Keep PRs focused — one feature or fix per PR
- Test on both Windows and macOS where possible (GitHub Actions covers automated builds)

---

## Support the Project

Excel Image Inserter is free and open-source. If it saves you time on real work, please consider supporting development — it helps cover maintenance, a Windows code-signing certificate, and future features.

Available channels: Telegram Stars, Buy Me a Coffee, GitHub Sponsors, crypto (TON / USDT), and Boosty. Details and addresses in **[DONATE.md](DONATE.md)**.

Non-monetary contributions help just as much: file useful bug reports, write docs, open PRs, or share the tool with someone who'd benefit.

---

## License

MIT — see [LICENSE](LICENSE) for the full text.

---

## Author

**Ilya Moskvin**
GitHub: [@moskva1988](https://github.com/moskva1988)
