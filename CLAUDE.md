# CLAUDE.md

Project context for Claude Code. Read this before making changes.

## What this is

**table2sheet**: an offline tool that extracts tables from hospital/provider price lists
(text PDFs, scanned PDFs/images, Excel/CSV) and outputs clean XLSX/CSV. Its main job is
**CPT preprocessing for Maxicare provider mapping**: it produces the 6-column Maxicare template
that a separate "CPT Mapping" tool consumes.

Template columns (`CPT_TEMPLATE_HEADER` in `table2sheet.py`):
`MAXICARE CPT CODE | MAXICARE CPT DESC | PROVIDER CODE | PROVIDER CPT DESC | DEPARTMENT/COST CENTER | RATES`
- The two Maxicare columns stay blank (filled later by the mapping tool).
- RATES prefers the OP/outpatient rate.
- A blank provider code is filled with the department name; a missing department is inferred from description keywords.
- Leading zeros in codes (e.g. `010000000`) must always be preserved as text.

## Layout

| Path | Role |
|---|---|
| `table2sheet.py` | Core: extraction engines (pdfplumber for text PDFs, Tesseract OCR for scans, openpyxl for sheets), CPT preprocessing (`preprocess_cpt`, `_schema_union_merge`, `build_cpt_mapping_table`), CLI, and Tk GUI |
| `web_server.py` | Browser UI at `http://localhost:8765` (standard library only). Endpoints: `/api/convert`, `/api/download/...`, `/api/ping`, `/api/shutdown`. Uploads/downloads go to `scratch/` |
| `run.bat`, `restart.bat`, `launch.ps1` | Windows launchers; prefer the bundled `runtime\python`, then fall back to system Python |
| `create_shortcut.ps1` | Creates the `Run table2sheet` shortcut for the folder it runs in (shortcuts hold absolute paths, so they're made per laptop and `*.lnk` is gitignored) |
| `build_runtime.py` | Builds `runtime\python` (Python + packages) and `runtime\tesseract` (copied from an installed Tesseract) |
| `requirements.txt` | pdfplumber, openpyxl, numpy, opencv-python-headless, Pillow, pytesseract |
| `samples/` | **Developer dataset only** (provider price lists, matched/expected outputs). Not deployed, not used at runtime |
| `README.md`, `USER_GUIDE.md`, `START-HERE.txt` | Docs: developer/CLI, full user guide, end-user quick note |

## Working rules

- **The developer's local folder is the source of truth.** The repo started as an upload of an
  older copy. If newer local files arrive (e.g. on a `local-latest` branch), merge them in:
  keep their code changes, but keep the repo's layout (`samples/`, `requirements.txt`,
  `build_runtime.py`, `create_shortcut.ps1`, docs). Don't refactor `table2sheet.py` or
  `web_server.py` until that merge is done.
- **Target users are on locked-down Windows company laptops**: no admin rights, no installers,
  and `.bat` double-click may be blocked (hence the PowerShell shortcut). Keep everything
  offline and runnable from the bundled `runtime\`. Don't add dependencies that need
  admin rights or internet at runtime.
- **Deployment** is done by the developer: build `runtime\`, share the folder (without
  `samples\`) via Google Drive, and run `create_shortcut.ps1` on each laptop. Users' results
  stay on their own laptops.
- Never commit `runtime/`, `scratch/`, `*.lnk`, or user output files.
- When changing CLI options or behavior, update `README.md` and `USER_GUIDE.md` to match.
- `samples/` files are real provider rate data; don't send them to external services.

## Known follow-ups (samples, deferred)

- `samples/PRICELIST 2026 ... .csv` is a noisy OCR-style export (garbled header/text), not a clean reference.
- `samples/test_cpt_mapping_output.csv` predates the "blank provider code -> department" rule.
- `samples/880000000012226_ST. LUKE'S ... MATCHED FILE.xlsx` is the best example of the expected end result.

## Quick checks

```bash
python -m pip install -r requirements.txt
python table2sheet.py "samples/Radiology Prices_MSMC - Christopher, Jr. Via.xlsx" --cpt -o out/
python table2sheet.py --web   # web UI on localhost:8765
```
