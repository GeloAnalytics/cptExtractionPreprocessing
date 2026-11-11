# table2sheet

Offline converter: **PDF or image tables -> XLSX / CSV** (built for CPT / price-list mapping).
No internet, no cloud API, no account. Everything runs on your machine.

For a full step-by-step guide with diagrams and troubleshooting tips, see the [USER_GUIDE.md](USER_GUIDE.md).

## Quick Start (No Installation / Zero Admin)
Just double-click:
```
run.bat
```
*(or double-click the `Run table2sheet` shortcut).*
The application opens immediately in your default web browser at `http://localhost:8765` (just like the CPT mapping tool). No installer, no administrator rights, and no command prompt needed (see `START-HERE.txt`).

## Install (optional / CLI use)
* Text-based PDFs (like Excel-printed price lists) need **nothing else**.
* Scanned PDFs / photos / screenshots need the free **Tesseract OCR** engine:
  * Windows: `winget install UB-Mannheim.TesseractOCR`
  * macOS: `brew install tesseract`   Linux: `apt install tesseract-ocr`
  * If it is not on PATH: `--tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"` or set env `TESSERACT_CMD`.
  * Other languages: install the language pack and use `--lang eng+fil`.

## Use
```
python table2sheet.py price_list.pdf                 # -> price_list.xlsx next to the PDF
python table2sheet.py price_list.pdf -f both -o out  # xlsx + csv into folder "out"
python table2sheet.py scan.png -f csv
python table2sheet.py page1.png page2.png --merge -o combined.xlsx
python table2sheet.py folder_of_scans --merge -o combined.xlsx
python table2sheet.py                                # no arguments = simple window (GUI)
```

| Option | Meaning |
|---|---|
| `-f xlsx\|csv\|both` | output format (default xlsx) |
| `-o PATH` | output folder (or file name with `--merge` or single input) |
| `--cpt` | CPT Preprocessing: extract Code, Description, Department, and OPD rates; auto-generates both master spreadsheet and official Maxicare CPT Mapping CSV (`_cpt_mapping.csv`) |
| `--opd-only` | in CPT mode, keep only OPD/OP rates (hide ward, room, and inpatient columns) |
| `--rate-col NAME` | isolate a specific Rate column in CPT mode (default: keeps all room rates) |
| `--sheet NAME` | sheet name when input is an Excel workbook (default: auto-detects) |
| `--merge` | all inputs become one table (multi-page scans as separate images) |
| `--ocr auto\|always\|never` | auto = OCR only for pages with no text layer |
| `--precise` | bordered scans: OCR every cell separately (slower, most accurate) |
| `--numbers` | convert prices to real numbers (codes stay text) |
| `--text-cols 0,2` | columns always kept as text (default: first column = code) |
| `--ncols N` / `--columns "A,B,C"` | force column count / set header names |
| `--no-header` | table has no header row |
| `--merge-wrapped` | glue wrapped description lines to the row above |
| `--keep-preamble` | keep title lines above the table |
| `--pages 1-3,7` | only some PDF pages |
| `--compare ref.csv` | diff the result against a reference CSV |

## How it works
* **Digital / Vector PDFs**: Uses a native vector table extraction engine that inspects exact cell boundaries, lines, and rects. Multi-line cells (such as wrapped email addresses or long street names) are preserved cleanly within their respective rows instead of splitting into separate phantom rows.
* **Scanned PDFs / images**: Local Tesseract OCR; images are deskewed, ruling lines are detected, and bordered tables are read row-by-row or cell-by-cell with `--precise`.
* **Number & Code Safety**: Leading zeros on phone numbers (`0917...`), area codes (`074...`), and item codes (`0100...`) are always preserved as text strings, preventing Excel from stripping zeros or converting to scientific notation (`9.18E+09`).

## Accuracy notes
* Text-layer PDFs: exact. Always still spot-check a few rows.
* OCR: ~96% of rows exactly right on a clean 300 dpi bordered scan with `--precise`
  (~75-85% without it). Best results: scan at 300 dpi or more, straight, high contrast.
  **Always review OCR output before using it for mapping**, especially codes and prices.

## Sharing between laptops (Google Drive)
**Option A - portable app, no Python on the other laptops (recommended)**
1. On ONE Windows PC with Python, double-click `build_windows.bat` (needs internet once; install Tesseract first to include scan/OCR support).
2. It creates `table2sheet_portable.zip` (Tesseract bundled, works fully offline). Upload it to Google Drive.
3. On any other Windows laptop: download, unzip, double-click `table2sheet.exe`.

**Option B - share this folder, install per laptop**
Share the `table2sheet_app` folder via Drive. On each laptop run `install_windows.bat` (or `./install.sh` on Mac/Linux) once; it makes a desktop shortcut.
Needs Python 3.10+ and internet only during that install.

The exe build is windowed (GUI). For command-line use, use Option B or `python table2sheet.py ...`.
