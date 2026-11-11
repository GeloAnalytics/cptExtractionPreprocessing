#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
table2sheet - fully offline PDF / image table -> XLSX / CSV converter.

* Text-layer PDFs  : words + coordinates are read directly (pdfplumber). No OCR, 100% exact text.
* Scanned PDFs / images : local Tesseract OCR (no internet, no cloud API).
* Columns are found automatically from vertical white-space "gutters" (or ruling lines in
  scanned images), per page, so shifted layouts between pages are handled.
* Cell text is kept exactly as printed by default (leading zeros in codes are preserved).

Usage:
    python table2sheet.py price_list.pdf                  # -> price_list.xlsx
    python table2sheet.py scan.png --format csv
    python table2sheet.py a.pdf b.pdf --format both -o out_folder
    python table2sheet.py page1.png page2.png --merge -o combined.xlsx
    python table2sheet.py                                  # opens the GUI
"""
from __future__ import annotations

import argparse
import csv
import math
import os
import re
import shutil
import statistics
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Sequence

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp", ".gif"}
PDF_EXTS = {".pdf"}
SHEET_EXTS = {".xlsx", ".xls", ".csv"}


# --------------------------------------------------------------------------- data types
@dataclass
class Word:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float
    conf: float = 100.0

    @property
    def xc(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def yc(self) -> float:
        return (self.top + self.bottom) / 2

    @property
    def h(self) -> float:
        return self.bottom - self.top


@dataclass
class Options:
    ocr: str = "auto"                 # auto | always | never
    dpi: int = 300                    # render resolution for OCR of PDF pages
    lang: str = "eng"                 # tesseract language(s), e.g. "eng+fil"
    min_conf: float = 30.0            # drop OCR words below this confidence
    ncols: Optional[int] = None       # force number of columns (default: auto)
    column_names: Optional[List[str]] = None
    header: bool = True               # first table row is a header
    keep_preamble: bool = False       # keep title lines above the table
    merge_wrapped: bool = False       # glue wrapped continuation lines onto the previous row
    key_col: int = 0                  # column that must be non-empty to start a new row
    numbers: bool = False             # convert numeric-looking cells to numbers (xlsx)
    text_cols: Sequence[int] = (0,)   # columns always kept as text (codes)
    gutter: float = 0.5               # min gutter width, in multiples of median text height
    tesseract_cmd: Optional[str] = None
    pages: Optional[str] = None       # e.g. "1-3,7"
    precise: bool = False             # bordered scans: OCR every cell separately (slower, more accurate)
    cpt: bool = False                 # CPT mode: 4 standard columns (Code, Description, Department, Rate) + deduplicate
    rate_col: Optional[str] = None    # column name / keyword for Rate in CPT mode (default: auto-detects OPD/Rate)
    opd_only: bool = False            # CPT mode: keep only OPD/OP rate column (suppress all ward/room/inpatient cols)
    sheet: Optional[str] = None       # sheet name when input is an Excel workbook (default: auto-detects)
    log: Callable[[str], None] = field(default=print, repr=False)


# --------------------------------------------------------------------------- helpers
def parse_pages(spec: Optional[str], n: int) -> List[int]:
    if not spec:
        return list(range(n))
    out: List[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out += range(int(a or 1) - 1, min(int(b or n), n))
        elif part:
            out.append(int(part) - 1)
    return list(dict.fromkeys(p for p in out if 0 <= p < n))


def setup_tesseract(cmd: Optional[str]):
    import pytesseract
    cmd = cmd or os.environ.get("TESSERACT_CMD")
    if not cmd:                      # Tesseract bundled next to the app (portable install)
        here = Path(getattr(sys, "_MEIPASS", "")) if getattr(sys, "frozen", False) else Path(__file__).parent
        bases = {here, Path(sys.executable).parent, Path(__file__).parent}
        for b in list(bases):
            bases.add(b.parent)
        for base in bases:
            for sub in ("tesseract", "runtime/tesseract", "../runtime/tesseract", "../tesseract"):
                for name in ("tesseract.exe", "tesseract"):
                    c = (base / sub / name).resolve()
                    if c.exists():
                        cmd = str(c)
                        os.environ.setdefault("TESSDATA_PREFIX", str(c.parent / "tessdata"))
                        break
                if cmd:
                    break
            if cmd:
                break
    if not cmd and not shutil.which("tesseract"):
        for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                  r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                  "/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract"):
            if os.path.exists(c):
                cmd = c
                break
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:  # noqa
        raise RuntimeError(
            "Tesseract OCR engine not found. It is only needed for scanned PDFs/images.\n"
            "Install it once (offline use afterwards): Windows: winget install UB-Mannheim.TesseractOCR | "
            "macOS: brew install tesseract | Linux: apt install tesseract-ocr\n"
            "Then set --tesseract-cmd or env TESSERACT_CMD if it is not on PATH."
        ) from e
    return pytesseract


# --------------------------------------------------------------------------- word acquisition
def words_from_text_layer(page) -> List[Word]:
    ws = page.extract_words(x_tolerance=1.5, y_tolerance=2, keep_blank_chars=False)
    return [Word(w["text"], w["x0"], w["x1"], w["top"], w["bottom"]) for w in ws if w["text"].strip()]


def _deskew_angle(gray) -> float:
    """Projection-profile skew estimate in degrees (searches +-4 deg)."""
    import cv2
    import numpy as np

    small = gray
    if max(gray.shape) > 1600:
        f = 1600 / max(gray.shape)
        small = cv2.resize(gray, None, fx=f, fy=f, interpolation=cv2.INTER_AREA)
    bw = (cv2.threshold(small, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1] > 0).astype(np.uint8)
    h, w = bw.shape
    best, best_ang = -1.0, 0.0
    for ang in np.arange(-4.0, 4.01, 0.25):
        M = cv2.getRotationMatrix2D((w / 2, h / 2), float(ang), 1.0)
        r = cv2.warpAffine(bw, M, (w, h), flags=cv2.INTER_NEAREST)
        score = float(np.var(r.sum(axis=1)))
        if score > best:
            best, best_ang = score, float(ang)
    return best_ang


def preprocess(img):
    """Grayscale, upscale small images, deskew. Returns (PIL image, scale_used)."""
    import cv2
    import numpy as np
    from PIL import Image

    arr = np.array(img.convert("L"))
    scale = 1.0
    if arr.shape[1] < 2400:
        scale = 2.0 if arr.shape[1] >= 1000 else 3.0
        arr = cv2.resize(arr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    ang = _deskew_angle(arr)
    if abs(ang) >= 0.25:
        h, w = arr.shape
        M = cv2.getRotationMatrix2D((w / 2, h / 2), ang, 1.0)
        arr = cv2.warpAffine(arr, M, (w, h), flags=cv2.INTER_CUBIC, borderValue=255)
    return Image.fromarray(arr), scale


def vertical_rules(img) -> List[float]:
    """x positions (px of img) of long vertical ruling lines (bordered tables)."""
    import cv2
    import numpy as np

    arr = np.array(img.convert("L"))
    bw = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    h, w = bw.shape
    kern = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, h // 6)))
    v = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kern)
    cols = np.where(v.sum(axis=0) > 0)[0]
    xs: List[float] = []
    for c in cols:                      # merge adjacent pixel columns into one line
        if xs and c - xs[-1][-1] <= 3:
            xs[-1].append(c)
        else:
            xs.append([c])
    return [float(np.mean(g)) for g in xs]


def remove_lines(img):
    """White-out long horizontal/vertical ruling lines so OCR sees only text."""
    import cv2
    import numpy as np
    from PIL import Image

    arr = np.array(img.convert("L"))
    bw = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    h, w = bw.shape
    hk = cv2.getStructuringElement(cv2.MORPH_RECT, (max(45, w // 50), 1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(45, h // 60)))
    vlines = cv2.dilate(cv2.morphologyEx(bw, cv2.MORPH_OPEN, vk), np.ones((1, 5), np.uint8))
    hlines = cv2.dilate(cv2.morphologyEx(bw, cv2.MORPH_OPEN, hk), np.ones((3, 1), np.uint8))
    lines = cv2.bitwise_or(vlines, hlines)
    arr = arr.copy()
    arr[lines > 0] = 255
    return Image.fromarray(arr)


_ARTIFACT = re.compile(r'^[|_=~\u2014\u2013\u2012\u00ac\u00b7.,:;\'"`*+]+$')


_EDGE = re.compile(r'^[|_=~\u2014\u2013]+|[|_=~]+$')


def _clean_token(t: str, conf: float, min_conf: float):
    """Strip ruling-line debris glued to a word; return None to drop the token."""
    t = (t or "").strip()
    if not t or _ARTIFACT.match(t):
        return None
    c = _EDGE.sub("", t)
    if c != t:                      # glued debris lowers confidence: trust the cleaned word
        return c or None
    return t if conf >= min_conf else None


def horizontal_rules(img) -> List[float]:
    """y positions (px) of long horizontal ruling lines (row separators of bordered tables)."""
    import cv2
    import numpy as np

    arr = np.array(img.convert("L"))
    bw = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    h, w = bw.shape
    kern = cv2.getStructuringElement(cv2.MORPH_RECT, (max(60, w // 4), 1))
    hz = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kern)
    rows = np.where(hz.sum(axis=1) > 0)[0]
    ys: List[List[int]] = []
    for r in rows:
        if ys and r - ys[-1][-1] <= 3:
            ys[-1].append(r)
        else:
            ys.append([r])
    return [float(np.mean(g)) for g in ys]


def _read(img, psm, opts, pytesseract, dx, dy, pad):
    from pytesseract import Output

    cfg = f"--oem 1 --psm {psm} -c preserve_interword_spaces=1"
    d = pytesseract.image_to_data(img, lang=opts.lang, config=cfg, output_type=Output.DICT)
    out = []
    for i, t in enumerate(d["text"]):
        try:
            conf = float(d["conf"][i])
        except ValueError:
            conf = -1
        t = _clean_token(t, conf, opts.min_conf)
        if t is None:
            continue
        x, y, w, h = d["left"][i] - pad + dx, d["top"][i] - pad + dy, d["width"][i], d["height"][i]
        out.append(Word(t, x, x + w, y, y + h, conf))
    return out


def _ocr_strips(clean, ys, opts: Options, pytesseract, vrules=()):
    """OCR each row strip (between consecutive horizontal rules); with opts.precise and
    vertical rules, each cell is read on its own."""
    import numpy as np
    from PIL import ImageOps

    heights = sorted(b - a for a, b in zip(ys, ys[1:]) if b - a > 12)
    med = heights[len(heights) // 2] if heights else 0
    words: List[Word] = []
    W = clean.size[0]
    xs = sorted(vrules)
    for a, b in zip(ys, ys[1:]):
        if b - a <= 12:
            continue
        top, bot = int(a) + 2, int(b) - 1
        psm = 7 if (b - a) <= 1.7 * med else 6
        if opts.precise and len(xs) >= 3:
            for x1, x2 in zip(xs, xs[1:]):
                cx1, cx2 = int(x1) + 3, int(x2) - 2
                if cx2 - cx1 < 8:
                    continue
                cell = clean.crop((cx1, top, cx2, bot))
                if (np.array(cell) < 128).sum() < 20:       # empty cell
                    continue
                cell = ImageOps.expand(cell, border=10, fill=255)
                words += _read(cell, psm, opts, pytesseract, cx1, top, 10)
        else:
            strip = ImageOps.expand(clean.crop((0, top, W, bot)), border=12, fill=255)
            words += _read(strip, psm, opts, pytesseract, 0, top, 12)
    return words


def words_from_ocr(img, opts: Options):
    """Returns (words, rule_xs) in the pixel space of the preprocessed image."""
    pytesseract = setup_tesseract(opts.tesseract_cmd)
    from pytesseract import Output

    pimg, _ = preprocess(img)
    rules = vertical_rules(pimg)
    hys = horizontal_rules(pimg)
    pimg = remove_lines(pimg)
    if len(hys) >= 4:                       # bordered table: read row by row
        return _ocr_strips(pimg, hys, opts, pytesseract, rules), rules
    cfg = "--oem 1 --psm 6 -c preserve_interword_spaces=1"
    d = pytesseract.image_to_data(pimg, lang=opts.lang, config=cfg, output_type=Output.DICT)
    words = []
    for i, t in enumerate(d["text"]):
        try:
            conf = float(d["conf"][i])
        except ValueError:
            conf = -1
        t = _clean_token(t, conf, opts.min_conf)
        if t is None:
            continue
        x, y, w, h = d["left"][i], d["top"][i], d["width"][i], d["height"][i]
        words.append(Word(t, x, x + w, y, y + h, conf))
    return words, rules


# --------------------------------------------------------------------------- layout analysis
def group_rows(words: List[Word]) -> List[List[Word]]:
    if not words:
        return []
    mh = statistics.median(w.h for w in words)
    tol = 0.55 * mh
    rows: List[List[Word]] = []
    centers: List[float] = []
    for w in sorted(words, key=lambda w: w.yc):
        if rows and abs(w.yc - centers[-1]) <= tol:
            rows[-1].append(w)
            centers[-1] = statistics.fmean(x.yc for x in rows[-1])
        else:
            rows.append([w])
            centers.append(w.yc)
    return [sorted(r, key=lambda w: w.x0) for r in rows]


def segments(row: List[Word], merge_gap: float):
    """Merge words separated by ordinary word-spacing into (x0, x1) segments."""
    segs = []
    for w in row:
        if segs and w.x0 - segs[-1][1] <= merge_gap:
            segs[-1][1] = max(segs[-1][1], w.x1)
        else:
            segs.append([w.x0, w.x1])
    return segs


def find_gutters(tab_rows, merge_gap, min_gutter, x_lo, x_hi, tolerance):
    """Return list of (center, width) of vertical white-space bands between columns."""
    n = int(math.ceil(x_hi - x_lo)) + 2
    cover = [0] * n
    for row in tab_rows:
        for a, b in segments(row, merge_gap):
            for i in range(max(0, int(a - x_lo)), min(n, int(b - x_lo) + 1)):
                cover[i] += 1
    gutters, start = [], None
    for i, c in enumerate(cover + [10 ** 9]):
        if c <= tolerance and start is None:
            start = i
        elif c > tolerance and start is not None:
            wdt = i - start
            if start > 0 and i < n and wdt >= min_gutter:
                gutters.append((x_lo + start + wdt / 2, float(wdt)))
            start = None
    return gutters


def detect_boundaries(rows, ncols, opts: Options, rules: Sequence[float] = ()):
    """Decide column boundary x positions for one page. Returns (boundaries, first_table_row_idx)."""
    allw = [w for r in rows for w in r]
    if not allw:
        return [], 0
    mh = statistics.median(w.h for w in allw)
    merge_gap = 0.5 * mh
    min_gutter = max(1.0, opts.gutter * mh)

    tab_idx = [i for i, r in enumerate(rows) if len(segments(r, merge_gap)) >= 2]
    if not tab_idx:
        return [], 0
    first = tab_idx[0]
    tab_rows = [rows[i] for i in tab_idx]
    x_lo = min(w.x0 for w in allw)
    x_hi = max(w.x1 for w in allw)

    if len(rules) >= 3 and (ncols is None or len(rules) - 1 == ncols):
        inner = sorted(rules)[1:-1]          # inner ruling lines = column separators
        return inner, first

    # The header row is often wider/centred than its data (e.g. "PRIVATE/2-BED"), so it is
    # left out of the gutter search; a small share of stray rows is tolerated.
    gut_rows = tab_rows[1:] if opts.header and len(tab_rows) > 3 else tab_rows
    cand = []
    for frac in ((0.02,) if ncols is None else (0.02, 0.05, 0.10, 0.20)):
        tol = max(1, int(frac * len(gut_rows)))
        cand = find_gutters(gut_rows, merge_gap, min_gutter, x_lo, x_hi, tol)
        if ncols is None or len(cand) >= ncols - 1:
            break
    if ncols is not None and len(cand) > ncols - 1:
        cand = sorted(sorted(cand, key=lambda g: -g[1])[: ncols - 1])
    return [c for c, _ in cand], first



def crossing(rows, boundaries) -> int:
    """Number of words that straddle a column boundary (0 for a clean fit)."""
    n = 0
    for r in rows:
        for w in r:
            for b in boundaries:
                if w.x0 + 0.5 < b < w.x1 - 0.5:
                    n += 1
                    break
    return n


def choose_boundaries(rows, own, first, good_sets, target, opts, rules):
    """Pick the column boundaries for a page: its own if clean, else a proven layout from another page."""
    body = rows[first:]
    nwords = max(1, sum(len(r) for r in body))
    ok_count = (target is None) or (len(own) + 1 == target)
    if ok_count and crossing(body, own) <= 0.002 * nwords:
        return own
    best, best_score = own, None
    if ok_count:
        best_score = crossing(body, own)
    for cand in good_sets:
        sc = crossing(body, cand)
        if best_score is None or sc < best_score:
            best, best_score = list(cand), sc
    return best


def assign(rows, boundaries) -> List[List[str]]:
    out = []
    nc = len(boundaries) + 1
    for r in rows:
        cells = [[] for _ in range(nc)]
        for w in r:
            k = sum(1 for b in boundaries if w.xc > b)
            cells[k].append(w.text)
        out.append([" ".join(c) for c in cells])
    return out


# --------------------------------------------------------------------------- page / file level
def extract_pages(path: Path, opts: Options):
    """Yield (page_label, rows_of_words, rule_xs) for every page/image."""
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        from PIL import Image, ImageSequence
        im = Image.open(path)
        frames = [f.copy() for f in ImageSequence.Iterator(im)]
        for i, f in enumerate(frames):
            if opts.ocr == "never":
                raise RuntimeError("Image input needs OCR; remove --ocr never")
            words, rules = words_from_ocr(f, opts)
            yield f"{path.name}:{i + 1}", words, rules
    elif ext in PDF_EXTS:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for pi in parse_pages(opts.pages, len(pdf.pages)):
                page = pdf.pages[pi]
                words = [] if opts.ocr == "always" else words_from_text_layer(page)
                rules: List[float] = []
                if opts.ocr != "never" and (opts.ocr == "always" or len(words) < 5):
                    img = page.to_image(resolution=opts.dpi).original
                    words, rules = words_from_ocr(img, opts)
                    opts.log(f"  page {pi + 1}: OCR")
                yield f"{path.name}:{pi + 1}", words, rules
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}")


def clean_cell(text: Optional[str]) -> str:
    """Clean cell contents, repairing wrapped emails, hyphenated words, and multi-line breaks."""
    if not text:
        return ""
    t = str(text)
    # Fix broken hyphenated words: e.g. "hospi-\ntal" -> "hospital"
    t = re.sub(r'(\w+)-\s*\n\s*(\w+)', r'\1\2', t)
    # Fix broken email addresses across lines: e.g. "user@domain.\ncom" or "user@\ngmail.com"
    t = re.sub(r'(@[\w.-]+)\s*\n\s*([\w.-]+)', r'\1\2', t)
    t = re.sub(r'([\w.-]+)\s*\n\s*(com|net|org|gov|ph|edu|mil)', r'\1\2', t)
    # Replace remaining newlines with spaces so single cells stay clean
    t = re.sub(r'\s*\n\s*', ' ', t).strip()
    return t


def extract_tables_from_pdf(path: Path, opts: Options) -> Optional[List[List[str]]]:
    """High-accuracy extraction using PDF vector lines, rects, and text layout."""
    import pdfplumber
    table: List[List[str]] = []
    header: Optional[List[str]] = None
    target_cols = opts.ncols

    try:
        with pdfplumber.open(path) as pdf:
            page_indices = parse_pages(opts.pages, len(pdf.pages))
            if not page_indices:
                return None

            for pi in page_indices:
                page = pdf.pages[pi]
                # 1. Try native vector line/rect table extraction first
                tabs = page.extract_tables()
                if not tabs:
                    # 2. Try text-clustering table extraction (for borderless tables)
                    tabs = page.extract_tables(table_settings={
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                    })

                if not tabs:
                    continue

                for tab in tabs:
                    if not tab:
                        continue
                    for ri, row in enumerate(tab):
                        cleaned = [clean_cell(c) for c in row]
                        if not any(cleaned):
                            continue

                        # Determine if this row is a header row
                        is_hdr = False
                        if opts.header:
                            row_norm = " ".join(norm(c) for c in cleaned)
                            if header is None and ri == 0:
                                is_hdr = True
                            elif header is not None:
                                same = sum(norm(a) == norm(c) for a, c in zip(cleaned, header)) / max(1, len(header))
                                if same >= 0.5:
                                    is_hdr = True
                                elif "name" in row_norm and ("facility" in row_norm or "hospital" in row_norm):
                                    is_hdr = True

                        if is_hdr:
                            if header is None:
                                header = cleaned
                            continue

                        table.append(cleaned)
    except Exception as e:
        opts.log(f"  Notice: Vector table engine ({e}), falling back to layout analysis")
        return None

    if not table:
        return None

    # Determine modal column count to align pages
    col_counts = [len(r) for r in table]
    if header:
        col_counts.append(len(header))
    max_c = target_cols or (statistics.mode(col_counts) if col_counts else None)

    padded = []
    for r in table:
        if max_c and len(r) != max_c:
            r = (r + [""] * max_c)[:max_c]
        padded.append(r)

    if header:
        if max_c and len(header) != max_c:
            header = (header + [""] * max_c)[:max_c]
        if opts.column_names:
            padded.insert(0, list(opts.column_names))
        else:
            padded.insert(0, header)
    elif opts.column_names:
        padded.insert(0, list(opts.column_names))

    return padded


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().lower()


def norm_s(val: str) -> str:
    if val is None:
        return ""
    return " ".join(str(val).strip().split())


def norm_rate(val: str) -> str:
    if val is None:
        return ""
    s = str(val).strip().replace(",", "")
    try:
        f = float(s)
        return f"{f:.2f}"
    except ValueError:
        return norm_s(val).upper()


def extract_table_from_spreadsheet(path: Path, opts: Options) -> List[List[str]]:
    """Read a table directly from an Excel (.xlsx, .xls) or CSV (.csv) file."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with open(path, newline="", encoding=enc) as f:
                    return [[str(c) if c is not None else "" for c in row] for row in csv.reader(f)]
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Unable to decode CSV file: {path}")

    # XLSX / XLS
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    target_ws = None

    if opts.sheet and opts.sheet in wb.sheetnames:
        target_ws = wb[opts.sheet]
        opts.log(f"  Using specified sheet: '{opts.sheet}'")
    else:
        # Score each sheet by non-empty rows and CPT column relevance
        best_sheet = None
        best_score = -1

        for name in wb.sheetnames:
            ws = wb[name]
            first_row = None
            row_count = 0
            for r in ws.iter_rows(values_only=True):
                if any(c is not None and str(c).strip() for c in r):
                    if first_row is None:
                        first_row = [str(c or "").strip() for c in r]
                    row_count += 1
            if not first_row or row_count == 0:
                continue

            norm_hdrs = [h.upper() for h in first_row]
            score = 0
            if any("CODE" in h for h in norm_hdrs): score += 2
            if any("DESC" in h or "DESCRIPTION" in h for h in norm_hdrs): score += 2
            if any("DEPT" in h or "DEPARTMENT" in h for h in norm_hdrs): score += 2
            if any("OPD" in h or "OP" in h or "RATE" in h for h in norm_hdrs): score += 2
            if name.lower() == "sheet1": score += 1
            if row_count > 100: score += 1

            if score > best_score:
                best_score = score
                best_sheet = name

        chosen = best_sheet or wb.sheetnames[0]
        opts.log(f"  Loaded spreadsheet sheet: '{chosen}'")
        target_ws = wb[chosen]

    rows = []
    for r in target_ws.iter_rows(values_only=True):
        rows.append([str(c) if c is not None else "" for c in r])
    wb.close()
    return rows


_RATE_KEYWORDS = [
    "OPD", "OP", "OUTPATIENT", "OUT-PATIENT", "OUT PATIENT",
    "WARD", "SUITE", "PRIVATE", "2-BED", "SEMI-PRIVATE", "SEMI PRIVATE",
    "ICU", "CCU", "NICU", "PICU",
    "RATE", "RATES", "PRICE", "PRICES", "FEE", "FEES", "AMOUNT", "CHARGE",
    "PLAIN", "CONTRAST", "PACKAGE", "INPATIENT", "INPATIENTS",
    "HMO", "PHILHEALTH", "MEDICARE",
    "CASH", "PAYING", "NON-HMO", "NON HMO", "STANDARD", "REGULAR", "TARIFF", "SRP",
    "NET", "GROSS", "DISCOUNT", "DISCOUNTED", "TOTAL CHARGE", "UNIT PRICE", "COST", "BASE RATE",
    "DELIVERY", "OPERATING ROOM", "RECOVERY ROOM", "ER", "EMERGENCY ROOM", "EMERGENCY",
    "CLASS A", "CLASS B", "CLASS C", "ROOM RATE", "PROFESSIONAL FEE", "PF",
]

_EXCLUDE_KEYWORDS = [
    "MAXICARE", "REMARK", "CONCERN", "COMMENT", "NOTE",
    "TAGGING", "TAG", "STATUS", "RESPONSE", "DATE", "STARTDATE", "ENDDATE",
    "HIERARCHY", "TYPECODE", "ROOMTYPECODE", "PROVIDERNAME", "PROVIDER NAME",
    "FLAG", "ACTION", "VERIFICATION", "CREATION", "COST CENTER", "DEPARTMENT", "DEPT"
]

# ---------------------------------------------------------------------------
# Department inference & section helpers
# ---------------------------------------------------------------------------
_DEPT_MAP = [
    ("RADIOLOGY", "Radiology"),
    ("LABORATORY", "Laboratory"),
    ("CT SCAN", "CT Scan"),
    ("ULTRASOUND", "Ultrasound"),
    ("X-RAY", "X-Ray"),
    ("XRAY", "X-Ray"),
    ("MRI", "MRI"),
    ("MAMMOGRAPHY", "Mammography"),
    ("NUCLEAR MEDICINE", "Nuclear Medicine"),
    ("CARDIOLOGY", "Cardiology"),
    ("ENDOSCOPY", "Endoscopy"),
    ("HEMODIALYSIS", "Hemodialysis"),
    ("PHARMACY", "Pharmacy"),
    ("PHYSICAL THERAPY", "Physical Therapy"),
    ("REHABILITATION", "Rehabilitation"),
    ("PULMONARY", "Pulmonary"),
    ("RESPIRATORY", "Respiratory"),
    ("OPHTHALMOLOGY", "Ophthalmology"),
    ("PATHOLOGY", "Pathology"),
    ("DENTAL", "Dental"),
    ("DERMATOLOGY", "Dermatology"),
    ("SURGERY", "Surgery"),
    ("OUTPATIENT", "Outpatient"),
    ("EMERGENCY", "Emergency"),
    ("ONCOLOGY", "Oncology"),
    ("GASTROENTEROLOGY", "Gastroenterology"),
    ("NEUROLOGY", "Neurology"),
    ("ENT", "ENT"),
    ("PEDIATRICS", "Pediatrics"),
    ("OBSTETRICS", "OB-Gynecology"),
    ("GYNECOLOGY", "OB-Gynecology"),
    ("OB-GYN", "OB-Gynecology"),
    ("ANESTHESIOLOGY", "Anesthesiology"),
    ("BLOOD BANK", "Blood Bank"),
]

# Regex to detect embedded codes in descriptions like "85025 - CBC", "85025 CBC", or "CT-101: BRAIN"
_EMBEDDED_CODE_RE = re.compile(
    r"^([A-Z0-9]{2,6}[-\.][A-Z0-9]{2,8}|[0-9]{4,6}|[A-Z]{2,4}[0-9]{3,6})(?:[-:—\s]+)(.+)$"
)
# Medical service description prefixes that must NOT be split into code + description.
# e.g. "UTZ-LOWER ABDOMEN" is a service name, not code="UTZ-LOWER", desc="ABDOMEN".
_MEDICAL_DESC_PREFIX_RE = re.compile(
    r"^(?:UTZ|UIZ|UTZZ|X-RAY|XRAY|CT\s|CT-|MRI|2D\s*ECHO|ECG|EKG|CBC|RBC|WBC|PLT|AFP|PSA|HBA1C|DRE)[-\s]",
    re.I
)

# Patterns for report headers, pagination, or disclaimers that should be dropped
_REPORT_METADATA_PATTERNS = [
    re.compile(r"^(PAGE\s+\d+|RUN\s+DATE|PRINTED\s+ON|AS\s+OF\s+DATE|REPORT\s+ID|PREPARED\s+BY|DATE\s+GENERATED|TOTAL\s+RECORDS|CONFIDENTIAL|ALL\s+RIGHTS\s+RESERVED)", re.I),
    re.compile(r"^(EFFECTIVE\s+DATE|APPROVED\s+BY|FACILITY|HOSPITAL|MEDICAL\s+CENTER|DISCLAIMER|NOTICE)", re.I),
    re.compile(r"^(DESCRIPTION|TEST\s+DESCRIPTION|PROCEDURE|PROCEDURES|PARTICULARS|SERVICE|SERVICES|EXAMINATION|NAME\s+OF\s+TEST|ITEM\s+DESCRIPTION|CPT\s+CODE|PROCEDURE\s+CODE)$", re.I),
]


def detect_section_department(text: str) -> Optional[str]:
    """If a row contains a department banner (e.g. 'DEPARTMENT OF RADIOLOGY' or 'CLINICAL LABORATORY'), return canonical name."""
    t = text.strip().upper()
    if not t or len(t) > 50:
        return None
    for k, name in _DEPT_MAP:
        if k in t:
            return name
    return None


def _is_metadata_or_repeated_header(desc: str, code: str, rates: List[str]) -> bool:
    """Return True if this row is report metadata, a repeated table header, or noise."""
    d = str(desc or "").strip()
    has_rates = any(bool(str(r).strip()) for r in rates)
    # Row with description matching repeated table header words is always junk
    d_u = d.upper()
    if d_u in ("DESCRIPTION", "TEST DESCRIPTION", "PROCEDURE", "PROCEDURES", "PARTICULARS", "SERVICE", "SERVICES", "EXAMINATION", "NAME OF TEST", "ITEM DESCRIPTION"):
        return True
    if not d:
        return not (code or has_rates)
    if has_rates and code:
        return False
    for pat in _REPORT_METADATA_PATTERNS:
        if pat.search(d):
            if not has_rates or pat.match(d):
                return True
    return False


def infer_dept_from_file(p: Path, text_sample: str = "") -> str:
    """Infer department name from PDF preamble text or filename stem."""
    # Check first few lines of preamble for an exact standalone dept keyword
    for line in text_sample.splitlines()[:20]:
        lc = line.strip().upper()
        for k, name in _DEPT_MAP:
            if lc == k or lc == f"DEPARTMENT OF {k}" or lc == f"{k} DEPARTMENT":
                return name
    # Check filename stem
    stem_u = p.stem.upper()
    for k, name in _DEPT_MAP:
        if k in stem_u:
            return name
    # Check full preamble text
    if text_sample:
        sample_u = text_sample.upper()
        for k, name in _DEPT_MAP:
            if k in sample_u:
                return name
    return ""


def _get_pdf_preamble(path: Path) -> str:
    """Return first-page text from a PDF (empty string for scanned PDFs)."""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            if pdf.pages:
                txt = pdf.pages[0].extract_text()
                return txt or ""
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
# OCR two-tier header flattener (for CT-scan style scanned tables)
# ---------------------------------------------------------------------------
_HMO_KEYWORDS = re.compile(
    r"\b(HMO|CHARGES?|MLANG|MEDICAL\s+CENTER|HOSPITAL|DIAGNOSTIC|SPECIALIST)\b", re.I
)
_TIER1_KEYWORDS = re.compile(
    r"\b(PLAIN|WITH\s+CONTRAST|CONTRAST|PACKAGE|PROCEDURE|PROCEDURES)\b", re.I
)
_TIER2_KEYWORDS = re.compile(
    r"\b(OPD|OUTPATIENT|INPATIENT|INPATIENTS|WARD|ROOM|SUITE|ICU|ER|EMERGENCY)\b", re.I
)


def clean_ocr_2tier_header(table: List[List[str]]) -> List[List[str]]:
    """Remove preamble title rows and merge 2-tier column headers into one row.

    For CT-scan-style tables (PLAIN / WITH CONTRAST / PACKAGE across top,
    OPD / INPATIENTS along the bottom of the header band), combines them
    into unambiguous names like ``PLAIN OPD``, ``WITH CONTRAST INPATIENTS``,
    ``PACKAGE (W/ CONTRAST & SYRINGE) OPD``, etc.
    """
    if not table or len(table) < 2:
        return table

    # Drop leading HMO / hospital preamble rows (≤2 non-empty cells, hospital keywords)
    while table:
        r0 = table[0]
        non_empty = [c for c in r0 if str(c).strip()]
        r0_str = " ".join(non_empty).upper()
        if len(r0) >= 4 and len(non_empty) <= 2 and _HMO_KEYWORDS.search(r0_str):
            table = table[1:]
        else:
            break

    if len(table) < 2:
        return table

    r0, r1 = table[0], table[1]
    r0_u = " ".join(str(c) for c in r0).upper()
    r1_u = " ".join(str(c) for c in r1).upper()

    is_tier1_row = bool(_TIER1_KEYWORDS.search(r0_u))
    is_tier2_row = bool(_TIER2_KEYWORDS.search(r1_u))

    if not (is_tier1_row and is_tier2_row):
        return table

    # Merge rows into a single combined header
    combined: List[str] = []
    for c0, c1 in zip(r0, r1):
        t0 = str(c0 or "").strip()
        t1 = str(c1 or "").strip()

        # Normalise tier-1 category name, stripping OCR noise
        t0_up = t0.upper()
        if "PACKAGE" in t0_up:
            cat = "PACKAGE (W/ CONTRAST & SYRINGE)"
        elif "CONTRAST" in t0_up:
            cat = "WITH CONTRAST"
        elif "PLAIN" in t0_up:
            cat = "PLAIN"
        elif "PROCEDURE" in t0_up:
            cat = "PROCEDURES"
        else:
            cat = re.sub(r"\s+", " ", t0).strip()

        # Normalise tier-2 sub-name
        t1_up = t1.upper()
        if "INPATIENT" in t1_up:
            sub = "INPATIENTS"
        elif "OPD" in t1_up or "OUTPATIENT" in t1_up:
            sub = "OPD"
        else:
            sub = re.sub(r"\s+", " ", t1).strip()

        if cat and sub and sub.upper() not in cat.upper():
            combined.append(f"{cat} {sub}")
        elif cat:
            combined.append(cat)
        else:
            combined.append(sub)

    return [combined] + table[2:]


# ---------------------------------------------------------------------------
# Bordered-table OCR with correct intra-strip word order
# ---------------------------------------------------------------------------
def _ocr_bordered_table(pimg, clean, hys: List[float], rules: List[float],
                        opts: "Options", pytesseract) -> List[List[str]]:
    """OCR a bordered scanned table row-by-row; return list-of-list-of-str rows
    (including the header row detected as row 0).

    Key improvement over the default _ocr_strips path: words within a strip are
    sorted by their *pixel* top (``orig_top``) so that wrapped text in the
    description column is reassembled in reading order, e.g.
    ``CT UPPER/LOWER ABDOMEN`` appears before ``TRIPHASIC`` in the same cell.
    """
    from PIL import ImageOps

    W = clean.size[0]
    xs = sorted(rules)

    words: List[Word] = []
    heights = sorted(b - a for a, b in zip(hys, hys[1:]) if b - a > 12)
    med = heights[len(heights) // 2] if heights else 0

    for a, b in zip(hys, hys[1:]):
        if b - a <= 12:
            continue
        top, bot = int(a) + 2, int(b) - 1
        # Use the strip centre as the nominal row y-coordinate for grouping
        strip_yc = (a + b) / 2.0
        psm = 7 if (b - a) <= 1.7 * med else 6

        strip = ImageOps.expand(clean.crop((0, top, W, bot)), border=12, fill=255)
        raw_words = _read(strip, psm, opts, pytesseract, 0, top, 12)

        for w in raw_words:
            # Store original pixel top; pin yc to strip centre for row grouping
            w_fixed = Word(w.text, w.x0, w.x1, strip_yc - 10, strip_yc + 10, w.conf)
            # pylint: disable=attribute-defined-outside-init
            w_fixed.orig_top = w.top  # type: ignore[attr-defined]
            words.append(w_fixed)

    if not words:
        return []

    rows = group_rows(words)
    b_cuts, first = detect_boundaries(rows, opts.ncols, opts, rules)

    nc = len(b_cuts) + 1
    result: List[List[str]] = []
    for r in rows:
        cells: List[List[Word]] = [[] for _ in range(nc)]
        for w in r:
            k = sum(1 for cut in b_cuts if w.xc > cut)
            cells[k].append(w)
        row_strs: List[str] = []
        for c in cells:
            # Sort by original pixel line first, then x, so wrapped lines are correct
            c_sorted = sorted(c, key=lambda x: (round(getattr(x, "orig_top", x.top) / 12), x.x0))
            row_strs.append(" ".join(x.text for x in c_sorted))
        result.append(row_strs)

    return result


# ---------------------------------------------------------------------------
# Multi-column scanned page detection & robust OCR
# ---------------------------------------------------------------------------
def detect_page_columns(pil_img) -> List[tuple]:
    """Detect independent side-by-side table column tracks in a scanned image.

    Returns a list of (x0, x1) pixel column ranges, e.g. 5 columns on Page 1 or 2 columns on Page 2.
    Returns [(0, w)] if the page is a single uniform table.
    """
    import cv2
    import numpy as np

    arr = np.array(pil_img.convert("L"))
    h, w = arr.shape
    bw = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    contours, _ = cv2.findContours(bw, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    boxes = []
    for c in contours:
        bx, by, bw_w, bw_h = cv2.boundingRect(c)
        # Avoid the giant outer table frame spanning > 55% of width
        if 80 <= bw_w <= int(w * 0.55) and bw_h >= 60:
            boxes.append((bx, bx + bw_w))

    if not boxes:
        return [(0, w)]

    boxes.sort()
    merged = []
    for x0, x1 in boxes:
        if not merged:
            merged.append([x0, x1])
        else:
            # Sub-columns in the same table touch with gap <= 4px
            if x0 <= merged[-1][1] + 4:
                merged[-1][1] = max(merged[-1][1], x1)
            else:
                merged.append([x0, x1])

    cols = [(x0, x1) for x0, x1 in merged if (x1 - x0) >= 200]
    return cols if len(cols) >= 2 else [(0, w)]


def _is_ocr_noise_text(text: str) -> bool:
    """Return True for short OCR fragments or broken tokens that should be discarded."""
    s = (text or "").strip()
    if not s:
        return True
    if re.fullmatch(r"[\W_]+", s):
        return True
    if "|" in s or "||" in s:
        return True
    if re.search(r"[\]\)]\s*$", s) and not re.search(r"\(", s):
        return True
    if re.search(r"(?:^|\s)[A-Za-z]\s*([|/\\]|[\)])\s*(?:[A-Za-z]\s*)*$", s):
        return True

    s_upper = s.upper()
    if any(phrase in s_upper for phrase in [
        "APPROVED", "PRICE SUBJECT", "OTHER TEST", "ADMINISTRATION ONLY",
        "MISCELLANEOUS", "VACCINE", "FEE", "FEE ARRANGE", "DENTAL CHECK-UP"
    ]):
        return True

    alpha = sum(ch.isalpha() for ch in s)
    allowed_keywords = r"(?:UTZ|X-RAY|AFP|CBC|ECG|EKG|MRI|CT|BPS|TSH|HCG|FBS|PSA|PYLORI|CULTURE|STOOL|SERUM|PROTEIN|CHOLESTEROL|ALT|AST|PAP|PCR|DENGUE|VDRL|SEROLOGY|GRA|THYROID|BILIRUBIN|PLATELET|HEMOGLOBIN|TROP|HBSAG|HCV|HBV|BETA|HIV|PREGNANCY|LIVER|KIDNEY|SCROTUM|GALLBLADDER|CHOLESTERASE|SEMEN|URIC|ALBUMIN|MICRAL|GRAM|FERRITIN|PHOSPHORUS|POTASSIUM|CALCIUM|CHLORIDE|CREATININE|TOXIOD)"
    if alpha <= 2 and not re.search(allowed_keywords, s_upper):
        return True
    if re.fullmatch(r"[A-Za-z]{1,3}(?:\s+[A-Za-z]{1,3}){0,2}", s) and not re.search(allowed_keywords, s_upper):
        return True

    tokens = re.findall(r"[A-Za-z]+", s)
    short_tokens = [t for t in tokens if len(t) <= 2 and t.upper() not in {"CT", "UTZ", "ECG", "EKG", "AFP", "CBC", "TSH", "PCR", "HCG", "FBS", "PSA", "ALT", "AST", "HBV", "HCV"}]
    if len(short_tokens) >= 2 and not re.search(allowed_keywords, s_upper):
        return True
    if not re.search(allowed_keywords, s_upper) and len(re.findall(r"[A-Za-z]+", s)) <= 1:
        return True
    return False


def clean_ocr_item(raw_line: str) -> tuple[str, str]:
    """Clean a raw line from OCR into (description, price).

    Normalizes OCR typos, strips ruling-line artifacts, and fixes common
    medical prefix noise (e.g. D-RAY -> X-RAY, furz -> UTZ).
    """
    line = (raw_line or "").strip()
    if not line or len(line) < 3:
        return "", ""

    # Strip leading/trailing border noise characters
    line = re.sub(r'^[|_=~\u2014\u2013\u2012\u00ac\u00b7.,:;\'"`*+<>{}\[\]\\/-]+', '', line)
    line = re.sub(r'[|_=~\u2014\u2013\u2012\u00ac\u00b7.,:;\'"`*+<>{}\[\]\\/-]+$', '', line).strip()

    # Fix OCR typos inside numbers before extraction: e.g. 3:30.00 -> 330.00, 445,90 -> 445.90
    line = re.sub(r'(\d+):(\d{2}\.\d{2})', r'\1\2', line)
    line = re.sub(r'(\d+),(\d{2})\b', r'\1.\2', line)

    # ---- Price search strategy ----
    # Pattern 1: price at end (after whitespace/noise separator)
    m = re.search(r'[\s|~_=«\-]+(?:[₱$sS«iI|\s_~\-]+)?(?:[A-Za-z]{0,3})?([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2}|[0-9]{2,5}\.[0-9]{2})\s*[|_~«\-]*$', line)
    if not m:
        # Pattern 2: integer price at end (e.g. 1,200 or 1200) with optional stray OCR letter prefix
        m = re.search(r'[\s|~_=«\-]+(?:[₱$sS«iI|\s_~\-]+)?(?:[A-Za-z]{0,3})?([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{3,5})\s*[|_~«\-]*$', line)
    if not m:
        # Pattern 3: price embedded anywhere with separator chars (handles OCR line wrapping)
        m = re.search(r'[\s=\-—_~«]{2,}(?:[A-Za-z]{0,3})?([0-9]{1,3}(?:,[0-9]{3})*\.[0-9]{2}|[0-9]{3,5})\s*$', line)
    if not m:
        # Pattern 4: price with only 1 decimal digit at end (e.g. 550.0)
        m = re.search(r'[\s=\-—_~«]+(?:[A-Za-z]{0,3})?([0-9]{2,5}\.[0-9])\s*$', line)
    if not m:
        # Pattern 5: price with space between parts (e.g. 385 00 -> 38500 is wrong; treat as 385.00)
        m2 = re.search(r'[\s=\-—_~«]+(?:[A-Za-z]{0,3})?([0-9]{3})\s+([0-9]{2})\s*$', line)
        if m2:
            # Reconstruct into decimal string for float conversion
            line = line[:m2.start()] + f' {m2.group(1)}.{m2.group(2)}'
            m = re.search(r'([0-9]{3}\.[0-9]{2})\s*$', line)

    price = ""
    desc = line
    if m:
        raw_p = m.group(1).replace(",", "")
        try:
            val = float(raw_p)
            # Sanity check: medical prices typically under 100000
            if 1.0 <= val <= 99999.0:
                price = f"{val:.2f}"
                desc = line[:m.start()].strip()
        except ValueError:
            pass

    # Clean description
    desc = re.sub(r'^[|_=~\u2014\u2013\u2012\u00ac\u00b7.,:;\'"`*+<>{}\[\]\\/-]+', '', desc)
    desc = re.sub(r'[|_=~\u2014\u2013\u2012\u00ac\u00b7.,:;\'"`*+<>{}\[\]\\/-]+$', '', desc).strip()
    desc = re.sub(r'[\s|_=~\u2014\u2013-]+$', '', desc).strip()

    # Strip trailing OCR noise (equals signs, dashes, single letters)
    desc = re.sub(r'[\s=\-—_~«sS«iI]{3,}$', '', desc).strip()
    # Strip trailing lone characters/symbols after valid words
    desc = re.sub(r'\s+[a-z]{1,2}$', '', desc).strip()

    # Fix X-RAY prefix noise
    desc = re.sub(r'^[A-Za-z0-9\W]{0,3}[-\s]?RAY(?=[A-Z\s\-])', 'X-RAY ', desc, flags=re.I)
    desc = re.sub(r'^[A-Za-z0-9\W]{0,3}[-\s]?RAY\b', 'X-RAY', desc, flags=re.I)
    desc = re.sub(r'^[Cc]RAY\b', 'X-RAY', desc, flags=re.I)

    # Fix UTZ prefix noise
    desc = re.sub(r'^[A-Za-z\W]{0,3}(?:UTZ|UIZ|URZ|UEZ|furz|lurz|jurz)[-\s]?', 'UTZ-', desc, flags=re.I)

    desc = re.sub(r'\s+', ' ', desc).strip()
    return desc, price


_DEPT_KEYWORDS = [
    # CT Scan (highest precedence)
    (["CT SCAN", "CT ", "CT-", "CAT SCAN", "TOMOGRAPHY"], "CT Scan"),
    # X-Ray (before general organ names)
    (["X-RAY", "XRAY", "-RAY", "CRAY", "CHEST PA", "APL", "AP/L", "SKELETAL", "SKULL SERIES", "WATERS VIEW", "TOWNE", "THORACIC", "LUMBOSACRAL", "PELVIS AP", "ANKLE AP", "FOOT AP", "HAND AP", "WRIST AP", "SPINE AP", "CERVICAL", "CLAVICLE", "RIBS", "SHOULDER AP", "FEMUR", "TIBIA", "FIBULA"], "X-Ray"),
    # Ultrasound (before general organ names)
    (["UTZ", "ULTRASOUND", "SONO", "DOPPLER", "DUPLEX", "ECHO", "BPS", "TRANSVAGINAL", "TRANSRECTAL", "PELVIS", "KUB", "THYROID UTZ", "BREAST UTZ"], "Ultrasound"),
    # Cardiology
    (["ECG", "EKG", "TREADMILL", "HOLTER", "STRESS TEST"], "Cardiology"),
    # Laboratory (comprehensive medical lab keywords)
    (["CBC", "HEMOGLOBIN", "PLATELET", "BLOOD", "URINE", "URINALYSIS", "STOOL", "FECALYSIS", "CHOLESTEROL", "GLUCOSE", "FBS", "HBA1C", "BUN", "CREATININE", "SGPT", "SGOT", "ALT", "AST", "LIPID", "SODIUM", "POTASSIUM", "TSH", "FT3", "FT4", "T3", "T4", "THYROID", "HEPATITIS", "HBSAG", "ANTI-HBS", "ANTI-HCV", "ANTI-HAV", "COVID", "SWAB", "ANTIGEN", "GRAM STAIN", "PAP SMEAR", "CULTURE", "BILIRUBIN", "TRIGLYCERIDES", "MICRAL", "SEMENALYSIS", "SEMMENALYSIS", "TYPHIDOT", "DENGUE", "VDRL", "AFP", "PSA", "CA-125", "CA 19-9", "BETA HCG", "TOTAL PROTEIN", "TOTALPROTEIN", "ALBUMIN", "CHLORIDE", "CHOLINESTERASE", "AMYLASE", "LIPASE", "SEROLOGY", "IMMUNOLOGY", "IMUNOLOGY", "HEMATOLOGY", "CHEMISTRY", "MICROSCOPY", "PARASITOLOGY", "CLINICAL", "DRUG TEST", "OCCULT BLOOD", "OCCULT"], "Laboratory"),
]


def infer_dept_from_desc(desc: str, fallback: str = "") -> str:
    """Infer medical department name from procedure description keywords."""
    d_u = (desc or "").upper()
    for keywords, dept in _DEPT_KEYWORDS:
        if any(k in d_u for k in keywords):
            return dept
    return fallback


def _ocr_multicolumn_page(pil_img, cols: List[tuple], opts: "Options", pytesseract) -> List[List[str]]:
    """OCR each independent column in a multi-column scanned page and stack rows vertically.

    Detects section headers, classifies departments, and formats rows as
    ['', description, department, price].
    """
    w, h = pil_img.size
    page_rows: List[List[str]] = []

    for ci, (cx0, cx1) in enumerate(cols):
        col_crop = pil_img.crop((max(0, cx0 - 4), 200, min(w, cx1 + 4), h - 80))
        pimg, _ = preprocess(col_crop)
        clean = remove_lines(pimg)
        txt = pytesseract.image_to_string(clean, config=f"--oem 1 --psm 6 -c preserve_interword_spaces=1", lang=opts.lang)
        lines = [l.strip() for l in txt.splitlines() if l.strip()]

        cur_dept = ""
        for l in lines:
            d, p = clean_ocr_item(l)
            if not d or len(d) < 2:
                continue
            if _is_ocr_noise_text(d):
                continue

            # Skip report metadata / signatures
            d_up = d.upper()
            if any(ign in d_up for ign in [
                "APPROVED BY", "APPROVED OF", "OPERATIONS OFFICER", "GRRICER",
                "EFFECTIVE:", "EFFECTIVE :", "PRICE SUBJECT TO CHANGE",
                "TOCHANGE", "CHIEF OPERATIONS", "CHIEF OP",
                "PREPARED BY", "NOTED BY", "DATE OF EFFECTIVITY",
            ]):
                continue

            # Skip pure-number or pure-symbol lines
            if re.match(r"^[0-9\W]+$", d):
                continue

            # Count meaningful alpha characters for noise detection
            alpha_chars = [c for c in d if c.isalpha()]
            alnum_chars = [c for c in d if c.isalnum()]
            alnum_ratio = len(alnum_chars) / max(1, len(d.replace(' ', '')))

            # Extra guardrail for OCR fragments that would otherwise survive a weak text metric.
            if re.search(r"\|", d):
                continue
            if re.search(r"(?:^|\s)[A-Za-z]\s*([|/\\]|[\)])\s*(?:[A-Za-z]\s*)*$", d):
                continue
            if re.search(r"[\]\)]\s*$", d) and not re.search(r"\(", d):
                continue
            if _is_ocr_noise_text(d):
                continue

            # Lines with no price: apply aggressive noise filter
            if not p:
                # Must have at least 4 alpha chars and 60% alnum ratio to keep
                if len(alpha_chars) < 4 or alnum_ratio < 0.60:
                    continue
                # Very short lines with no price: skip unless they look like a known test
                if len(d.strip()) <= 3:
                    continue

            # Check for section header row (e.g. ULTRASOUND, HEPATITIS PROFILE, etc.)
            if not p:
                inferred_hdr = infer_dept_from_desc(d)
                if inferred_hdr and len(d.split()) <= 4:
                    cur_dept = inferred_hdr
                    continue

            dept = cur_dept or infer_dept_from_desc(d)
            page_rows.append(["", d, dept, p])

    return page_rows


def preprocess_cpt(table: List[List[str]], opts: Options) -> List[List[str]]:
    """Standardize CPT tables: extract Code, Description, Department, and ALL rate columns
    (such as Suite, Ward, Private, OPD), excluding extraneous metadata columns,
    and deduplicate rows where (Description and/or Code, Department, Rates) match."""
    if not table:
        return table

    hdr = table[0]
    norm_hdrs = [str(h or "").strip().upper() for h in hdr]

    # 1. Detect Code column (avoiding Maxicare / external codes and rate-like columns)
    code_idx = None
    if "CODE" in norm_hdrs:
        code_idx = norm_hdrs.index("CODE")
    if code_idx is None:
        for p in ["PROVIDER CODE", "PROVIDER CPT CODE", "CPT CODE", "ITEM CODE", "PROCEDURE CODE"]:
            for i, h in enumerate(norm_hdrs):
                if p in h and not any(ex in h for ex in ["MAXICARE"]):
                    code_idx = i
                    break
            if code_idx is not None:
                break
    if code_idx is None:
        for i, h in enumerate(norm_hdrs):
            # Avoid "HMO", "PLAIN", "CONTRAST", etc. being treated as code column
            if ("CODE" in h and "DESC" not in h
                    and not any(ex in h for ex in ["MAXICARE"])
                    and not any(rk in h for rk in _RATE_KEYWORDS)):
                code_idx = i
                break

    # 2. Detect Description column (avoiding Maxicare desc)
    desc_idx = None
    if "DESCRIPTION" in norm_hdrs:
        desc_idx = norm_hdrs.index("DESCRIPTION")
    elif "DESC" in norm_hdrs:
        desc_idx = norm_hdrs.index("DESC")
    if desc_idx is None:
        for p in ["TEST DESCRIPTION", "PROVIDER CPT DESC", "PROVIDER DESC", "CPT DESC", "DESCRIPTION", "DESC", "PROCEDURE", "PARTICULARS", "SERVICE"]:
            for i, h in enumerate(norm_hdrs):
                if i != code_idx and p in h and not any(ex in h for ex in ["MAXICARE"]):
                    desc_idx = i
                    break
            if desc_idx is not None:
                break

    # 3. Detect Department column
    dept_idx = None
    for p in ["DEPARTMENT", "DEPART", "DEPT", "COST CENTER", "SECTION", "UNIT"]:
        for i, h in enumerate(norm_hdrs):
            if i not in (code_idx, desc_idx) and p in h:
                dept_idx = i
                break
        if dept_idx is not None:
            break

    def norm_rate(v: str) -> str:
        s = re.sub(r"^(PHP|Php|php|PhP|₱|P|\$|EUR)\s*", "", str(v).strip())
        s = re.sub(r"\s+", "", s).replace(",", "")
        try:
            return f"{float(s):.2f}"
        except ValueError:
            return s

    # 4. Detect ALL Rate columns (Suite, Private, Ward, OPD, Rate, etc.)
    rate_indices: List[int] = []
    if opts.rate_col:
        hint_u = opts.rate_col.strip().upper()
        for i, h in enumerate(norm_hdrs):
            if i not in (code_idx, desc_idx, dept_idx) and (hint_u == h or hint_u in h):
                rate_indices.append(i)
                break

    if not rate_indices:
        for i, h in enumerate(norm_hdrs):
            if i in (code_idx, desc_idx, dept_idx):
                continue
            if any(ex in h for ex in _EXCLUDE_KEYWORDS):
                continue
            words = re.findall(r'[A-Za-z0-9\-]+', h)
            matched = any(rk in words or rk in h for rk in _RATE_KEYWORDS)
            if matched:
                rate_indices.append(i)
            elif len(table) > 1:
                # Check if data in column is predominantly numeric price/currency values
                sample_vals = [
                    re.sub(r"^(PHP|Php|php|PhP|₱|P|\$|EUR)\s*", "", str(r[i]).strip()).replace(",", "").replace(" ", "")
                    for r in table[1:50] if len(r) > i and r[i] is not None and str(r[i]).strip()
                ]
                if sample_vals and sum(bool(_NUM.match(v)) for v in sample_vals) / len(sample_vals) >= 0.75:
                    rate_indices.append(i)

    # Fallback to last column if no rate column could be identified
    if not rate_indices and len(hdr) > 0:
        for i in range(len(hdr) - 1, -1, -1):
            if i not in (code_idx, desc_idx, dept_idx):
                rate_indices.append(i)
                break

    # Smart statistical detection for missing core columns:
    core_indices = [code_idx, desc_idx, dept_idx]
    unassigned = [i for i in range(len(hdr)) if i not in core_indices and i not in rate_indices]

    # 1. Code: check unassigned columns for short alphanumeric tokens (<= 14 chars, no spaces)
    if code_idx is None and unassigned:
        for idx in list(unassigned):
            sample = [str(r[idx]).strip() for r in table[1:50] if len(r) > idx and r[idx] is not None and str(r[idx]).strip()]
            if sample:
                code_ratio = sum(bool(re.match(r"^[A-Za-z0-9\-_./]{2,14}$", v)) for v in sample) / len(sample)
                avg_len = sum(len(v) for v in sample) / len(sample)
                if code_ratio >= 0.7 and avg_len <= 12:
                    code_idx = idx
                    unassigned.remove(idx)
                    break

    # 2. Description: check unassigned columns for highest average text length
    if desc_idx is None and unassigned:
        best_desc, best_len = None, -1
        for idx in unassigned:
            sample = [str(r[idx]).strip() for r in table[1:50] if len(r) > idx and r[idx] is not None and str(r[idx]).strip()]
            if sample:
                avg_len = sum(len(v) for v in sample) / len(sample)
                if avg_len > best_len:
                    best_len = avg_len
                    best_desc = idx
        if best_desc is not None:
            desc_idx = best_desc
            unassigned.remove(best_desc)

    # 3. Department: check unassigned columns for matching known department names
    if dept_idx is None and unassigned:
        for idx in list(unassigned):
            sample = [str(r[idx]).strip().upper() for r in table[1:50] if len(r) > idx and r[idx] is not None and str(r[idx]).strip()]
            if sample and any(any(k in v for k, _ in _DEPT_MAP) for v in sample):
                dept_idx = idx
                unassigned.remove(idx)
                break

    c_idx, d_idx, dep_idx = code_idx, desc_idx, dept_idx

    rate_names = [hdr[i] for i in rate_indices]
    c_str = f"'{hdr[c_idx]}' (col {c_idx+1})" if c_idx is not None else "(not present)"
    d_str = f"'{hdr[d_idx]}' (col {d_idx+1})" if d_idx is not None else "(not present)"
    dep_str = f"'{hdr[dep_idx]}' (col {dep_idx+1})" if dep_idx is not None else "(not present)"
    opts.log(f"  CPT columns mapped: Code={c_str}, Description={d_str}, Department={dep_str}, Rates={rate_names}")

    out_table = [["Code", "Description", "Department"] + rate_names]

    seen_all = set()
    seen_desc_dept_rate = set()
    seen_code_dept_rate = set()

    dup_count = 0
    current_dept = ""

    for r in table[1:]:
        c_val = str(r[c_idx]).strip() if c_idx is not None and c_idx < len(r) and r[c_idx] is not None else ""
        d_val = str(r[d_idx]).strip() if d_idx is not None and d_idx < len(r) and r[d_idx] is not None else ""
        dep_val = str(r[dep_idx]).strip() if dep_idx is not None and dep_idx < len(r) and r[dep_idx] is not None else ""
        rate_vals = [str(r[ri]).strip() if ri < len(r) and r[ri] is not None else "" for ri in rate_indices]

        if not any([c_val, d_val, dep_val] + rate_vals):
            continue

        # In-table Section Department Banners (e.g. 'DEPARTMENT OF RADIOLOGY', 'CLINICAL LABORATORY')
        if not c_val and not any(bool(v.strip()) for v in rate_vals):
            sec_dept = detect_section_department(d_val)
            if sec_dept:
                current_dept = sec_dept
                continue  # section divider row consumed

        # Department forward-fill
        if dep_val:
            current_dept = dep_val
        elif current_dept:
            dep_val = current_dept

        # Drop report metadata or repeated headers across page breaks
        if _is_metadata_or_repeated_header(d_val, c_val, rate_vals):
            continue

        # Auto-extract embedded codes from description if Code column was absent.
        # Guard: skip if the description starts with a known medical service prefix
        # (e.g. UTZ-LOWER ABDOMEN, X-RAY KNEE APL) which should stay as full descriptions.
        if not c_val and d_val and not _MEDICAL_DESC_PREFIX_RE.match(d_val):
            m_code = _EMBEDDED_CODE_RE.match(d_val)
            if m_code:
                c_val = m_code.group(1).strip()
                d_val = m_code.group(2).strip()

        c_norm = norm_s(c_val).upper()
        d_norm = norm_s(d_val).upper()
        dep_norm = norm_s(dep_val).upper()
        rate_tuple = tuple(norm_rate(v) for v in rate_vals)

        is_dup = False
        k_all = (c_norm, d_norm, dep_norm, rate_tuple)
        k_desc = (d_norm, dep_norm, rate_tuple) if (d_norm and dep_norm) else (d_norm, rate_tuple) if d_norm else None
        k_code = (c_norm, dep_norm, rate_tuple) if (c_norm and dep_norm) else (c_norm, rate_tuple) if c_norm else None

        if k_all in seen_all:
            is_dup = True
        elif k_desc and k_desc in seen_desc_dept_rate:
            is_dup = True
        elif k_code and k_code in seen_code_dept_rate:
            is_dup = True

        if is_dup:
            dup_count += 1
            continue

        seen_all.add(k_all)
        if k_desc:
            seen_desc_dept_rate.add(k_desc)
        if k_code:
            seen_code_dept_rate.add(k_code)

        out_table.append([c_val, d_val, dep_val] + rate_vals)

    opts.log(f"  CPT Preprocessing: pruned {dup_count} duplicate rows; kept {len(out_table)-1} unique rows.")
    return out_table


def extract_single_doc(p: Path, opts: Options) -> Optional[List[List[str]]]:
    """Extract table from a single document using the optimal engine for its format.

    For scanned PDFs with bordered tables (e.g. CT-scan rate lists), uses a
    strip-aware OCR path that correctly sorts intra-cell wrapped words, then
    collapses any 2-tier column-header rows into one merged header row.
    """
    ext = p.suffix.lower()
    if ext in SHEET_EXTS:
        return extract_table_from_spreadsheet(p, opts)

    # 1. High-accuracy digital vector engine for text PDFs (exact text, zero OCR noise)
    if ext in PDF_EXTS and opts.ocr != "always":
        try:
            t = extract_tables_from_pdf(p, opts)
            if t:
                opts.log(f"  {p.name}: high-accuracy digital table extracted "
                         f"({len(t) - (1 if opts.header else 0)} rows, {len(t[0]) if t else 0} cols)")
                return t
        except Exception as ex:
            opts.log(f"  {p.name}: vector engine notice ({ex}), trying OCR...")

    # 2. Layout / OCR pipeline (for scanned PDFs, photos, or images)
    try:
        # --- Bordered-table fast path (horizontal rules detected) ---------------
        # For scanned PDFs that have clear ruled rows we use a strip-based OCR
        # that respects the original pixel order within each cell, then collapse
        # any two-tier column headers (e.g. PLAIN / OPD | PLAIN / INPATIENTS).
        if ext in PDF_EXTS:
            import pdfplumber
            with pdfplumber.open(p) as pdf:
                page_indices = parse_pages(opts.pages, len(pdf.pages))
                all_extracted: List[List[str]] = []
                master_hdr: Optional[List[str]] = None
                processed_any = False

                for pi in page_indices:
                    pil_img = pdf.pages[pi].to_image(resolution=opts.dpi).original
                    pytesseract = setup_tesseract(opts.tesseract_cmd)

                    # 1. Multi-column scanned table detection (e.g. 5 columns or 2 columns)
                    cols = detect_page_columns(pil_img)
                    if len(cols) >= 2:
                        processed_any = True
                        opts.log(f"  page {pi + 1}: multi-column scanned table ({len(cols)} columns)")
                        col_rows = _ocr_multicolumn_page(pil_img, cols, opts, pytesseract)
                        if col_rows:
                            if not all_extracted:
                                all_extracted = [["Code", "Description", "Department", "OP"]]
                                master_hdr = all_extracted[0]
                            all_extracted.extend(col_rows)
                        continue

                    # 2. Bordered table detection (horizontal ruling lines)
                    pimg, _ = preprocess(pil_img)
                    rules = vertical_rules(pimg)
                    hys = horizontal_rules(pimg)
                    if len(hys) >= 4:
                        processed_any = True
                        opts.log(f"  page {pi + 1}: OCR (bordered, strip mode)")
                        clean = remove_lines(pimg)
                        rows_raw = _ocr_bordered_table(pimg, clean, hys, rules, opts, pytesseract)
                        if rows_raw:
                            rows_clean = clean_ocr_2tier_header(rows_raw)
                            rows_clean = _drop_junk_rows(rows_clean, header_row=True)
                            if not all_extracted:
                                all_extracted = rows_clean
                                master_hdr = rows_clean[0] if rows_clean else None
                            else:
                                if rows_clean:
                                    start_idx = 1 if (master_hdr and len(rows_clean) > 1 and rows_clean[0] == master_hdr) else 0
                                    all_extracted.extend(rows_clean[start_idx:])
                        continue

                if processed_any and all_extracted:
                    opts.log(f"  {p.name}: OCR table extracted "
                             f"({len(all_extracted) - 1} rows, {len(all_extracted[0]) if all_extracted else 0} cols)")
                    return all_extracted
        elif ext in IMAGE_EXTS:
            from PIL import Image
            pil_img = Image.open(p).convert("RGB")
            pytesseract = setup_tesseract(opts.tesseract_cmd)
            cols = detect_page_columns(pil_img)
            if len(cols) >= 2:
                opts.log(f"  {p.name}: multi-column scanned table ({len(cols)} columns)")
                col_rows = _ocr_multicolumn_page(pil_img, cols, opts, pytesseract)
                if col_rows:
                    table_rows = [["Code", "Description", "Department", "OP"]] + col_rows
                    opts.log(f"  {p.name}: OCR table extracted "
                             f"({len(table_rows) - 1} rows, {len(table_rows[0])} cols)")
                    return table_rows
            pimg, _ = preprocess(pil_img)
            rules = vertical_rules(pimg)
            hys = horizontal_rules(pimg)
            if len(hys) >= 4:
                opts.log(f"  {p.name}: OCR (bordered, strip mode)")
                clean = remove_lines(pimg)
                rows_raw = _ocr_bordered_table(pimg, clean, hys, rules, opts, pytesseract)
                if rows_raw:
                    rows_clean = clean_ocr_2tier_header(rows_raw)
                    rows_clean = _drop_junk_rows(rows_clean, header_row=True)
                    opts.log(f"  {p.name}: OCR table extracted "
                             f"({len(rows_clean) - 1} rows, {len(rows_clean[0]) if rows_clean else 0} cols)")
                    return rows_clean

        # --- Generic OCR path --------------------------------------------------
        pages = []
        for label, words, rules in extract_pages(p, opts):
            if not words:
                opts.log(f"  {label}: no text detected")
                continue
            pages.append((label, group_rows(words), rules))

        detected = []
        for label, rows, rules in pages:
            b, first = detect_boundaries(rows, opts.ncols, opts, rules)
            detected.append((label, rows, rules, b, first))

        counts = [len(b) + 1 for _, _, _, b, _ in detected if b]
        target = opts.ncols or (statistics.mode(counts) if counts else None)

        good_sets, seen = [], set()
        for _, _, _, b, _ in detected:
            if b and target and len(b) + 1 == target:
                k = tuple(round(x) for x in b)
                if k not in seen:
                    seen.add(k)
                    good_sets.append(b)

        table: List[List[str]] = []
        header: Optional[List[str]] = None
        for label, rows, rules, b, first in detected:
            if not b:
                opts.log(f"  {label}: no table structure detected, skipped")
                continue
            if target and len(b) + 1 != target:
                b, first = detect_boundaries(rows, target, opts, rules)
            b = choose_boundaries(rows, b, first, good_sets, target, opts, rules)
            body = rows if opts.keep_preamble else rows[first:]
            cells = assign(body, b)
            if opts.header and cells:
                hdr = cells[0]
                if header is None:
                    header = hdr
                    cells = cells[1:]
                else:
                    same = sum(norm(a) == norm(c) for a, c in zip(hdr, header)) / max(1, len(header))
                    if same >= 0.6:
                        cells = cells[1:]
            table.extend(_pad(cells, target))
        table = fix_overflow(table)
        if opts.merge_wrapped:
            table = _merge_wrapped(table, opts.key_col)
        if header is not None:
            header = _pad([header], target)[0]
            table.insert(0, header)
        if opts.column_names:
            names = list(opts.column_names)
            if opts.header and header is not None:
                table[0] = names
            else:
                table.insert(0, names)
        if table:
            opts.log(f"  {p.name}: OCR table extracted "
                     f"({len(table) - (1 if opts.header else 0)} rows, {len(table[0]) if table else 0} cols)")
        return table
    except Exception as ex:
        opts.log(f"  ⚠️ {p.name}: Extraction failed ({ex})")
        return None


# ---------------------------------------------------------------------------
# Junk-row filter
# ---------------------------------------------------------------------------
_JUNK_PATTERN = re.compile(
    r"^[^A-Za-z0-9]*$"           # only punctuation / spaces
    r"|^[a-z]{1,3}(\s+[a-z]{1,5})*$"  # very short all-lowercase tokens (OCR noise)
)
# Matches OCR word-wrap artifacts like "tial oe ABDOMEN" — starts with lowercase
# fragment(s) before an uppercase word (e.g. residual end of a wrapped line)
_JUNK_WRAPPED_PATTERN = re.compile(r"^[a-z][a-z\s]{0,12}[A-Z]")


def _is_junk_row(row: List[str]) -> bool:
    """Return True if a row is an OCR artefact with no meaningful content."""
    non_empty = [str(c).strip() for c in row if str(c).strip()]
    if not non_empty:
        return True
    # Only 1 value and it looks like noise
    if len(non_empty) == 1:
        v = non_empty[0]
        # Lone lowercase word or very short garbage
        if _JUNK_PATTERN.match(v):
            return True
        # Lone all-lowercase multi-word string shorter than a real description
        if v == v.lower() and len(v) < 6:
            return True
    # First cell is an OCR word-wrap artifact (lowercase prefix + uppercase word)
    if non_empty:
        first = non_empty[0]
        if _JUNK_WRAPPED_PATTERN.match(first):
            return True
    return False


def _drop_junk_rows(table: List[List[str]], header_row: bool = True) -> List[List[str]]:
    """Remove blank and OCR-noise rows from a table, preserving the header row."""
    if not table:
        return table
    start = 1 if header_row else 0
    return table[:start] + [r for r in table[start:] if not _is_junk_row(r)]


# ---------------------------------------------------------------------------
# Schema-aligned multi-file merge
# ---------------------------------------------------------------------------
# CT Scan PLAIN OPD/INPATIENTS → canonical column names
# In opd_only mode: PLAIN OPD → "OP", WITH CONTRAST OPD → "OP (With Contrast)",
#                   PACKAGE OPD → "OP (Package)"
# In all-rates mode: same plain-OPD alias still merges into "OP"
_CT_PLAIN_OPD_ALIASES = {"PLAIN OPD", "PLAIN OP"}
_CT_OPD_REMAP = {
    # norm → (canonical_norm, display_name)
    "PLAIN OPD":  ("OP", "OP"),
    "PLAIN OP":   ("OP", "OP"),
    "WITH CONTRAST OPD":                            ("OP (WITH CONTRAST)", "OP (With Contrast)"),
    "PACKAGE OPD":                                  ("OP (PACKAGE)", "OP (Package)"),
    "PACKAGE (W/ CONTRAST & SYRINGE) OPD":          ("OP (PACKAGE)", "OP (Package)"),
    "PACKAGE (W/CONTRAST & SYRINGE) OPD":           ("OP (PACKAGE)", "OP (Package)"),
    "PACKAGE W/ CONTRAST & SYRINGE OPD":            ("OP (PACKAGE)", "OP (Package)"),
}

# Rate column names that are considered "OPD only" (kept when opd_only=True)
def _is_opd_rate_col(norm: str) -> bool:
    """Return True if this rate column should be kept in OPD-only mode."""
    n = norm.upper()
    # Always keep CT-scan OPD remapped columns
    if n in ("OP", "OP (WITH CONTRAST)", "OP (PACKAGE)"):
        return True
    # Keep columns that contain OPD or are exactly OP, but not inpatient/ward/room cols
    has_opd = "OPD" in n or n in ("OP", "O.P.", "OUT-PATIENT", "OUTPATIENT")
    is_inpatient = any(k in n for k in ("INPATIENT", "WARD", " IP", "SUITE", "PRIVATE",
                                         "PRESIDENTIAL", "ICU", "EMERGENCY", "DELIVERY"))
    return has_opd and not is_inpatient

def _schema_union_merge(doc_tables: List[tuple], opd_only: bool = False) -> List[List[str]]:
    """Merge multiple CPT tables (each already preprocessed to Code/Description/Department/rates)
    by aligning on shared column names rather than blindly concatenating rows.

    ``doc_tables`` is a list of ``(dept_name, table)`` pairs where ``table[0]``
    is the header row ``['Code', 'Description', 'Department', rate1, rate2, …]``.

    When ``opd_only`` is True, only OPD-keyed columns are kept. CT Scan columns are
    remapped: PLAIN OPD → OP, WITH CONTRAST OPD → OP (With Contrast),
    PACKAGE OPD → OP (Package).
    """
    def _ct_remap(raw: str):
        """Return (canonical_norm, display) after applying CT OPD remapping, or None if no remap."""
        norm = " ".join(str(raw).strip().split()).upper()
        return _CT_OPD_REMAP.get(norm)

    def _normalise(raw: str) -> str:
        r = _ct_remap(raw)
        return r[0] if r else " ".join(str(raw).strip().split()).upper()

    def _display(raw: str) -> str:
        r = _ct_remap(raw)
        return r[1] if r else str(raw).strip()

    # Build the union of all rate-column names, preserving first-seen order
    master_rate_norm: List[str] = []
    master_rate_display: List[str] = []
    for _, t in doc_tables:
        if not t or len(t[0]) <= 3:
            continue
        for col in t[0][3:]:
            col_norm = _normalise(col)
            if opd_only and not _is_opd_rate_col(col_norm):
                continue
            if col_norm not in master_rate_norm:
                master_rate_norm.append(col_norm)
                master_rate_display.append(_display(col))

    master_header = ["Code", "Description", "Department"] + master_rate_display
    merged: List[List[str]] = [master_header]

    for dept, t in doc_tables:
        if not t or len(t) <= 1:
            continue
        hdr = t[0]
        # Build column-index → master-index map (skip columns not in master when opd_only)
        col_map: dict = {0: 0, 1: 1, 2: 2}
        for i, col in enumerate(hdr[3:], start=3):
            col_norm = _normalise(col)
            if col_norm in master_rate_norm:
                col_map[i] = 3 + master_rate_norm.index(col_norm)

        for row in t[1:]:
            new_row: List[str] = [""] * len(master_header)
            for src_i, val in enumerate(row):
                if src_i in col_map:
                    new_row[col_map[src_i]] = val
            # Ensure Department is filled from the inferred dept name
            if not new_row[2] and dept:
                new_row[2] = dept
            merged.append(new_row)

    return merged


def convert_files(paths: Sequence[Path], opts: Options) -> List[List[str]]:
    """Extract and combine tables across all given input files.

    When CPT mode is active and multiple files are given, each file is
    preprocessed individually (so column detection works per-file), then the
    resulting CPT tables are merged by *column name* rather than by position.
    This avoids the classic column-shift bug where a Laboratory HMO column
    appears under Radiology's OP column just because they are both in position 2.

    Department is auto-populated from the filename / PDF preamble when no
    Department column is present in the extracted table.
    """
    if not opts.cpt or len(paths) <= 1:
        # ── Original single-file / non-CPT path ──────────────────────────────
        all_tables: List[List[str]] = []
        master_header: Optional[List[str]] = None
        for p in paths:
            t = extract_single_doc(p, opts)
            if not t:
                continue
            if opts.header and t:
                hdr, data = t[0], t[1:]
                if master_header is None:
                    master_header = hdr
                all_tables.extend(data)
            else:
                all_tables.extend(t)
        if master_header is not None:
            all_tables.insert(0, master_header)
        if opts.cpt and all_tables:
            all_tables = preprocess_cpt(all_tables, opts)
        return all_tables

    # ── Multi-file CPT path: per-file preprocessing + schema-union merge ─────
    doc_tables: List[tuple] = []
    for p in paths:
        opts.log(f"  Extracting {p.name} ...")
        t = extract_single_doc(p, opts)
        if not t or len(t) <= 1:
            continue

        # Infer department from preamble / filename before preprocessing
        preamble = _get_pdf_preamble(p) if p.suffix.lower() in PDF_EXTS else ""
        dept = infer_dept_from_file(p, preamble)

        # Per-file CPT preprocessing (column detection, deduplication)
        t_cpt = preprocess_cpt(t, opts)

        # Backfill Department column (index 2) from inferred dept
        if dept:
            for row in t_cpt[1:]:
                if len(row) > 2 and not str(row[2]).strip():
                    row[2] = dept

        doc_tables.append((dept, t_cpt))

    if not doc_tables:
        return []

    merged = _schema_union_merge(doc_tables, opd_only=opts.opd_only)
    opts.log(f"  → Merged {len(doc_tables)} files: {len(merged) - 1} total rows, "
             f"{len(merged[0])} columns")
    return merged



def fix_overflow(rows: List[List[str]]) -> List[List[str]]:
    """Text from a long left cell can spill into a numeric column; move it back to the left."""
    if not rows:
        return rows
    nc = len(rows[0])
    numeric = []
    for ci in range(nc):
        vals = [r[ci] for r in rows if len(r) > ci and r[ci].strip()]
        numeric.append(bool(vals) and sum(bool(_NUM.match(v.strip())) for v in vals) / len(vals) >= 0.9)
    for r in rows:
        for ci in range(min(nc, len(r))):
            if not numeric[ci] or not r[ci].strip():
                continue
            toks = r[ci].split()
            lead = []
            while toks and not _NUM.match(toks[0]):
                lead.append(toks.pop(0))
            if lead:
                left = next((k for k in range(ci - 1, -1, -1) if not numeric[k]), None)
                if left is not None:
                    r[left] = (r[left] + " " + " ".join(lead)).strip()
                    r[ci] = " ".join(toks)
    return rows


def _pad(rows, n):
    if not n:
        return rows
    return [(r + [""] * n)[:n] if len(r) != n else r for r in rows]


def _merge_wrapped(rows, key):
    out: List[List[str]] = []
    for r in rows:
        if out and key < len(r) and not r[key].strip():
            for i, c in enumerate(r):
                if c.strip() and i < len(out[-1]):
                    out[-1][i] = (out[-1][i] + " " + c).strip()
        else:
            out.append(list(r))
    return out


# --------------------------------------------------------------------------- writers
# Matches integers, decimals, negatives and accounting-style (1,234.56) or (1234)
_NUM = re.compile(r"^-?\(?[\d,]*\.?\d+\)?$")


def to_number(s: str):
    t = s.strip()
    if not t:
        return s
    # Strip common currency symbol prefixes (PHP, Php, ₱, P, $, EUR, etc.)
    t_clean = re.sub(r"^(PHP|Php|php|PhP|₱|P|\$|EUR)\s*", "", t)
    # Fix OCR-introduced spaces inside numbers: "1 ,267.50" -> "1,267.50"
    # and "1 3,260.00" -> "13,260.00" (digit-space-digit with comma pattern)
    if re.match(r"^-?\(?[\d\s,]*\.?\d+\)?$", t_clean):
        t_clean = re.sub(r"\s+", "", t_clean)  # remove all whitespace
    if not _NUM.match(t_clean):
        return s
    neg = (t_clean.startswith("(") and t_clean.endswith(")")) or t_clean.startswith("-")
    raw = t_clean.strip("()-").replace(",", "")
    # Keep codes with leading zeros (e.g. 01000, 007, 0917...) as text
    if len(raw) > 1 and raw.startswith("0") and not raw.startswith("0."):
        return s
    # Keep 10+ digit integers (e.g. phone numbers 9177997644) as text to avoid Excel scientific notation
    if len(raw) >= 10 and "." not in raw:
        return s
    try:
        v = float(raw)
    except ValueError:
        return s
    v = -v if neg else v
    return int(v) if v.is_integer() and "." not in raw else v


def write_csv(table, path: Path, opts: Options):
    rows = table
    if opts.numbers:
        # When a header row exists, keep row 0 as text; otherwise convert from row 0.
        data_start = 1 if opts.header and table else 0
        converted = [
            [c if i in opts.text_cols else _csv_num(c) for i, c in enumerate(r)]
            for r in table[data_start:]
        ]
        rows = table[:data_start] + converted
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerows(rows)


def _csv_num(c):
    v = to_number(c)
    return c if isinstance(v, str) else v


def write_xlsx(table, path: Path, opts: Options):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Table"
    for ri, r in enumerate(table, 1):
        for ci, v in enumerate(r, 1):
            # Skip numeric conversion on row 1 only when it's a header row
            is_header_row = (ri == 1 and opts.header)
            if not is_header_row and opts.numbers and (ci - 1) not in opts.text_cols:
                v = to_number(v)
            c = ws.cell(row=ri, column=ci, value=v)
            if isinstance(v, str):
                c.number_format = "@"            # keep codes as text (leading zeros)
            elif isinstance(v, (float, int)):
                c.number_format = "#,##0.00"
    if opts.header and table:
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    for ci in range(1, (len(table[0]) if table else 0) + 1):
        width = max(len(str(r[ci - 1])) for r in table[:500] if len(r) >= ci) + 2
        ws.column_dimensions[get_column_letter(ci)].width = min(max(width, 8), 60)
    wb.save(path)


# ---------------------------------------------------------------------------
# Maxicare CPT Mapping Tool Exporter
# ---------------------------------------------------------------------------
CPT_TEMPLATE_HEADER = [
    "MAXICARE CPT CODE",
    "MAXICARE CPT DESC",
    "PROVIDER CODE",
    "PROVIDER CPT DESC",
    "DEPARTMENT/COST CENTER",
    " RATES ",
]

def build_cpt_mapping_table(table: List[List[Any]]) -> List[List[str]]:
    """Transform an extracted table into the official Maxicare CPT Mapping template format.
    
    Columns:
    - MAXICARE CPT CODE: (blank, to be filled by the CPT mapping tool)
    - MAXICARE CPT DESC: (blank, to be filled by the CPT mapping tool)
    - PROVIDER CODE: procedure code, or blank if none
    - PROVIDER CPT DESC: procedure description
    - DEPARTMENT/COST CENTER: department name, or blank if none
    - RATES: outpatient rate (OP), or primary rate
    """
    if not table:
        return [CPT_TEMPLATE_HEADER]

    header = [str(c).strip() for c in table[0]]
    hdr_upper = [c.upper() for c in header]

    def _find_col(candidates):
        for cand in candidates:
            cand_u = cand.upper()
            for i, h in enumerate(hdr_upper):
                if h == cand_u or cand_u in h:
                    return i
        return None

    code_idx = _find_col(["PROVIDER CODE", "CODE", "ITEM CODE", "CPT CODE"])
    desc_idx = _find_col(["PROVIDER CPT DESC", "DESCRIPTION", "TEST DESCRIPTION", "PROCEDURE", "NAME", "ITEM"])
    dept_idx = _find_col(["DEPARTMENT/COST CENTER", "DEPARTMENT", "DEPT", "COST CENTER", "SECTION"])
    # Rate specifically prioritizes OP / OPD outpatient rates
    op_idx = _find_col(["OP", "OPD", "OUTPATIENT", "OUT-PATIENT", "PLAIN OPD", "RATES", "RATE", "PRICE"])

    if desc_idx is None:
        desc_idx = 1 if len(header) > 1 else 0

    out_rows = [CPT_TEMPLATE_HEADER]
    for r in table[1:]:
        if not r:
            continue
        c_val = r[code_idx] if code_idx is not None and code_idx < len(r) else ""
        d_val = r[desc_idx] if desc_idx is not None and desc_idx < len(r) else ""
        dept_val = r[dept_idx] if dept_idx is not None and dept_idx < len(r) else ""
        op_val = r[op_idx] if op_idx is not None and op_idx < len(r) else ""

        # Clean description
        d_str = "" if d_val is None else str(d_val).strip()
        if not d_str:
            continue
        if d_str.upper() in ("DESCRIPTION", "TEST DESCRIPTION", "PROCEDURE", "ITEM DESCRIPTION"):
            continue

        # Clean department: if missing, infer from description keywords
        dept_str = "" if dept_val is None or str(dept_val).strip().lower() in ("none", "nan", "null") else str(dept_val).strip()
        if not dept_str:
            dept_str = infer_dept_from_desc(d_str, fallback="General")

        # Clean code: if missing, use what is indicated in department
        c_str = "" if c_val is None or str(c_val).strip().lower() in ("none", "nan", "null") else str(c_val).strip()
        if not c_str:
            c_str = dept_str

        # Clean rate
        if op_val is None or op_val == "":
            rate_str = ""
        elif isinstance(op_val, (int, float)):
            rate_str = f"{float(op_val):.2f}"
        else:
            raw = str(op_val).strip()
            clean_num = re.sub(r"[^\d.]", "", raw)
            try:
                rate_str = f"{float(clean_num):.2f}"
            except ValueError:
                rate_str = raw

        out_rows.append(["", "", c_str, d_str, dept_str, rate_str])

    return out_rows


def build_tool_ready_table(table: List[List[Any]]) -> List[List[str]]:
    """Transform extracted table into clean 4-column Tool Ready table:
    [CPT Code, Description, Department, Rate]
    - Department is inferred from Description keywords if missing
    - CPT Code takes what is indicated in Department if missing
    - Rate is the Outpatient (OP) rate
    """
    cpt_mapping = build_cpt_mapping_table(table)
    out = [["CPT Code", "Description", "Department", "Rate"]]
    for r in cpt_mapping[1:]:
        out.append([r[2], r[3], r[4], r[5]])
    return out


def export_tool_ready_csv(table: List[List[Any]], dest: Path) -> Path:
    """Export table in clean Tool Ready format (CPT Code, Description, Department, Rate)."""
    tr_rows = build_tool_ready_table(table)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerows(tr_rows)
    return dest


def export_cpt_mapping_csv(table: List[List[Any]], dest: Path) -> Path:
    """Export table strictly formatted for the Maxicare CPT Mapping Tool."""
    cpt_rows = build_cpt_mapping_table(table)
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerows(cpt_rows)
    return dest


def compare_with_reference(table, ref_path: Path, log=print):
    with open(ref_path, newline="", encoding="utf-8-sig") as f:
        ref = list(csv.reader(f))
    n = max(len(ref), len(table))
    bad = 0
    shown = 0
    for i in range(n):
        a = ref[i] if i < len(ref) else None
        b = table[i] if i < len(table) else None
        if a != b:
            bad += 1
            if shown < 10:
                log(f"  row {i + 1}: expected {a} got {b}")
                shown += 1
    log(f"Compare: {len(ref)} reference rows, {len(table)} extracted rows, {bad} differing rows")
    return bad


# --------------------------------------------------------------------------- driver
def expand_inputs(items: Iterable[str]) -> List[Path]:
    out: List[Path] = []
    for it in items:
        p = Path(it)
        if p.is_dir():
            out += sorted(x for x in p.iterdir() if x.suffix.lower() in IMAGE_EXTS | PDF_EXTS | SHEET_EXTS)
        elif p.exists():
            out.append(p)
        else:
            raise FileNotFoundError(it)
    return out


def run(inputs: Sequence[str], fmt="xlsx", output: Optional[str] = None, merge=False,
        compare: Optional[str] = None, opts: Optional[Options] = None) -> List[Path]:
    opts = opts or Options()
    files = expand_inputs(inputs)
    if not files:
        raise ValueError("No input files found in the given paths.")
    # In CPT mode with multiple files, always merge so schema-union logic runs
    effective_merge = merge or (opts.cpt and len(files) > 1)
    jobs = [(files, files[0].stem)] if effective_merge else [([f], f.stem) for f in files]
    written: List[Path] = []
    for fl, stem in jobs:
        opts.log(f"Reading {', '.join(f.name for f in fl)} ...")
        table = convert_files(fl, opts)
        opts.log(f"  -> {max(0, len(table) - (1 if opts.header else 0))} data rows, "
                 f"{len(table[0]) if table else 0} columns")
        if compare:
            compare_with_reference(table, Path(compare), opts.log)
        outdir, outstem = fl[0].parent, stem
        if output:
            o = Path(output)
            if (effective_merge or len(jobs) == 1) and o.suffix.lower() in (".xlsx", ".csv"):
                outdir, outstem = o.parent, o.stem
            else:
                outdir = o
        outdir.mkdir(parents=True, exist_ok=True)
        for kind in (["xlsx", "csv"] if fmt == "both" else [fmt]):
            dest = outdir / f"{outstem}.{kind}"
            (write_xlsx if kind == "xlsx" else write_csv)(table, dest, opts)
            opts.log(f"  saved {dest} (Extraction: all duplicates & rates)")
            written.append(dest)
        if opts.cpt:
            tool_ready_dest = outdir / f"{outstem}_tool_ready.csv"
            export_tool_ready_csv(table, tool_ready_dest)
            opts.log(f"  saved {tool_ready_dest} (Tool Ready: CPT Code, Desc, Dept, Rate)")
            written.append(tool_ready_dest)

            cpt_dest = outdir / f"{outstem}_cpt_mapping.csv"
            export_cpt_mapping_csv(table, cpt_dest)
            opts.log(f"  saved {cpt_dest} (Maxicare CPT Mapping format)")
            written.append(cpt_dest)
    return written


def build_parser():
    p = argparse.ArgumentParser(description="Offline PDF/image table -> XLSX/CSV converter.")
    p.add_argument("inputs", nargs="*", help="PDF/image files or folders")
    p.add_argument("-f", "--format", choices=["xlsx", "csv", "both"], default="xlsx")
    p.add_argument("-o", "--output", help="output folder (or file name with --merge)")
    p.add_argument("--merge", action="store_true", help="combine all inputs into one table")
    p.add_argument("--ocr", choices=["auto", "always", "never"], default="auto")
    p.add_argument("--lang", default="eng", help="tesseract language, e.g. eng+fil")
    p.add_argument("--dpi", type=int, default=300)
    p.add_argument("--min-conf", type=float, default=30.0)
    p.add_argument("--pages", help="PDF pages, e.g. 1-3,7")
    p.add_argument("--ncols", type=int, help="force number of columns")
    p.add_argument("--columns", help="comma-separated column names to use as header")
    p.add_argument("--no-header", action="store_true")
    p.add_argument("--keep-preamble", action="store_true", help="keep title lines above the table")
    p.add_argument("--merge-wrapped", action="store_true", help="join wrapped lines to previous row")
    p.add_argument("--key-col", type=int, default=0)
    p.add_argument("--numbers", action="store_true", help="convert numeric cells to numbers")
    p.add_argument("--text-cols", default="0", help="0-based columns kept as text, e.g. 0,2")
    p.add_argument("--gutter", type=float, default=0.5, help="min column gap (x text height)")
    p.add_argument("--precise", action="store_true", help="bordered scans: OCR each cell separately (slower, most accurate)")
    p.add_argument("--cpt", action="store_true", help="CPT Mode: extract Code, Description, Department, and all room rates (Suite, Ward, OPD, etc.) and remove duplicates")
    p.add_argument("--opd-only", action="store_true", help="CPT Mode: keep only OPD/OP rates (hide ward, room, and inpatient columns)")
    p.add_argument("--rate-col", help="isolate a specific Rate column name/keyword in CPT mode (default: keeps all rates)")
    p.add_argument("--sheet", help="sheet name when input is an Excel workbook (default: auto-detects)")
    p.add_argument("--tesseract-cmd")
    p.add_argument("--compare", help="reference CSV to diff the result against")
    p.add_argument("--web", action="store_true", help="launch browser-based web interface (default when no inputs)")
    p.add_argument("--port", type=int, default=8765, help="port for local web interface (default: 8765)")
    p.add_argument("--gui", action="store_true", help="launch desktop Tkinter GUI window")
    return p


def main(argv=None):
    a = build_parser().parse_args(argv)
    if a.gui:
        return launch_gui()
    if a.web or not a.inputs:
        try:
            from web_server import start_server
        except ImportError:
            from .web_server import start_server
        return start_server(port=a.port or 8765)
    opts = Options(
        ocr=a.ocr, dpi=a.dpi, lang=a.lang, min_conf=a.min_conf, ncols=a.ncols,
        column_names=[c.strip() for c in a.columns.split(",")] if a.columns else None,
        header=not a.no_header, keep_preamble=a.keep_preamble, merge_wrapped=a.merge_wrapped,
        key_col=a.key_col, numbers=a.numbers,
        text_cols=tuple(int(x) for x in a.text_cols.split(",") if x != ""),
        gutter=a.gutter, tesseract_cmd=a.tesseract_cmd, pages=a.pages, precise=a.precise,
        cpt=a.cpt, rate_col=a.rate_col, opd_only=a.opd_only, sheet=a.sheet,
    )
    try:
        run(a.inputs, a.format, a.output, a.merge, a.compare, opts)
    except (ValueError, FileNotFoundError) as e:
        sys.exit(f"Error: {e}")


# --------------------------------------------------------------------------- GUI
def launch_gui():
    import os
    import subprocess
    import threading
    import time
    import tkinter as tk
    from pathlib import Path
    from tkinter import filedialog, messagebox, ttk

    # ── Modern Color Palette (2026 Fluent / Clean Slate) ──────────────────────
    BG             = "#F8FAFC"   # Slate 50 window background
    CARD_BG        = "#FFFFFF"   # Pure white card background
    BORDER         = "#E2E8F0"   # Slate 200 crisp flat border
    PRIMARY        = "#4F46E5"   # Indigo 600 primary action
    PRIMARY_HOVER  = "#4338CA"   # Indigo 700 hover state
    TEXT_MAIN      = "#0F172A"   # Slate 900 primary text
    TEXT_MUTED     = "#64748B"   # Slate 500 secondary / hint text
    TEXT_LIGHT     = "#94A3B8"   # Slate 400 faint label text
    CONSOLE_BG     = "#0B132B"   # Deep dark slate console
    CONSOLE_BAR    = "#1C2541"   # Console title bar
    CONSOLE_FG     = "#E2E8F0"   # Console standard text
    TAG_INFO       = "#38BDF8"   # Sky blue
    TAG_OK         = "#4ADE80"   # Mint green
    TAG_WARN       = "#FBBF24"   # Amber
    TAG_ERR        = "#F87171"   # Coral red

    # ── Root Window ───────────────────────────────────────────────────────────
    root = tk.Tk()
    root.title("table2sheet • PDF & Image Table Extractor")
    root.geometry("920x680")
    root.minsize(820, 580)
    root.configure(bg=BG)

    try:
        root.tk.call("tk", "scaling", 1.25)
    except Exception:  # noqa
        pass

    # Create clean modern in-memory window icon (replaces 1990s feather)
    try:
        ico = tk.PhotoImage(width=32, height=32)
        for x in range(32):
            for y in range(32):
                if 4 <= x <= 27 and 4 <= y <= 27:
                    ico.put("#4F46E5", (x, y))
                if 8 <= x <= 23 and 8 <= y <= 11:
                    ico.put("#818CF8", (x, y))
                if 8 <= x <= 23 and 15 <= y <= 17:
                    ico.put("#C7D2FE", (x, y))
                if 8 <= x <= 23 and 20 <= y <= 22:
                    ico.put("#C7D2FE", (x, y))
        root.iconphoto(False, ico)
    except Exception:  # noqa
        pass

    # ── ttk Base Theme Setup ──────────────────────────────────────────────────
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=BG, foreground=TEXT_MAIN, font=("Segoe UI", 10))
    style.configure("TFrame", background=BG)
    style.configure("TProgressbar", troughcolor="#E2E8F0", background=PRIMARY, thickness=5)
    style.configure("TCheckbutton", background=CARD_BG, foreground=TEXT_MAIN, font=("Segoe UI", 9))

    # ── Reactive State Variables ──────────────────────────────────────────────
    files: List[str] = []
    fmt_var    = tk.StringVar(value="xlsx")
    ocr_var    = tk.StringVar(value="auto")
    numbers    = tk.BooleanVar(value=True)
    merge_var  = tk.BooleanVar(value=False)
    precise    = tk.BooleanVar(value=False)
    outdir     = tk.StringVar(value="")
    lang_var   = tk.StringVar(value="eng")
    pages_var  = tk.StringVar(value="")
    ncols_var  = tk.StringVar(value="")
    cpt_var    = tk.BooleanVar(value=False)
    opd_only   = tk.BooleanVar(value=False)
    rate_var   = tk.StringVar(value="")
    sheet_var  = tk.StringVar(value="")
    status_msg = tk.StringVar(value="Ready • Add files to start")
    running    = threading.Event()
    last_dest  = tk.StringVar(value="")

    # ── Main Two-Column Layout ────────────────────────────────────────────────
    root.columnconfigure(0, weight=5)  # Left controls
    root.columnconfigure(1, weight=6)  # Right activity console
    root.rowconfigure(1, weight=1)

    # ── Top Modern Header Bar ─────────────────────────────────────────────────
    header_bar = tk.Frame(root, bg=BG, padx=18, pady=12)
    header_bar.grid(row=0, column=0, columnspan=2, sticky="ew")
    header_bar.columnconfigure(0, weight=1)

    title_box = tk.Frame(header_bar, bg=BG)
    title_box.pack(side="left")

    logo_badge = tk.Label(title_box, text=" table2sheet ", font=("Segoe UI", 14, "bold"),
                          bg=PRIMARY, fg="#FFFFFF", padx=6, pady=2)
    logo_badge.pack(side="left", padx=(0, 10))

    sub_box = tk.Frame(title_box, bg=BG)
    sub_box.pack(side="left")
    tk.Label(sub_box, text="Offline PDF & Image Table to Excel Converter",
             font=("Segoe UI", 11, "bold"), bg=BG, fg=TEXT_MAIN).pack(anchor="w")
    tk.Label(sub_box, text="100% private • No cloud APIs • Zero external installation required",
             font=("Segoe UI", 9), bg=BG, fg=TEXT_MUTED).pack(anchor="w")

    # Offline ready pill badge
    ready_pill = tk.Label(header_bar, text=" ● Engine Ready (Offline) ", font=("Segoe UI", 9, "bold"),
                          bg="#ECFDF5", fg="#059669", padx=10, pady=4,
                          highlightbackground="#A7F3D0", highlightthickness=1)
    ready_pill.pack(side="right")

    # ── Left Column Container (Scrollable or Clean Packed) ────────────────────
    left_col = tk.Frame(root, bg=BG, padx=16, pady=4)
    left_col.grid(row=1, column=0, sticky="nsew")
    left_col.columnconfigure(0, weight=1)

    # Helper: Create Flat Modern Card
    def create_card(parent):
        return tk.Frame(parent, bg=CARD_BG, highlightbackground=BORDER,
                        highlightthickness=1, bd=0, padx=14, pady=12)

    # ── CARD 1: INPUT FILES ───────────────────────────────────────────────────
    card_files = create_card(left_col)
    card_files.pack(fill="x", pady=(0, 10))
    card_files.columnconfigure(0, weight=1)

    card_files_header = tk.Frame(card_files, bg=CARD_BG)
    card_files_header.pack(fill="x", pady=(0, 8))
    tk.Label(card_files_header, text="Input Documents", font=("Segoe UI", 10, "bold"),
             bg=CARD_BG, fg=TEXT_MAIN).pack(side="left")
    files_count_lbl = tk.Label(card_files_header, text="0 files selected",
                               font=("Segoe UI", 9), bg="#F1F5F9", fg=TEXT_MUTED,
                               padx=8, pady=2)
    files_count_lbl.pack(side="right")

    # Clean listbox container
    lst_box_frame = tk.Frame(card_files, bg="#F8FAFC", highlightbackground=BORDER,
                             highlightthickness=1, bd=0)
    lst_box_frame.pack(fill="x", pady=(0, 8))

    file_listbox = tk.Listbox(lst_box_frame, height=4, bg="#F8FAFC", fg=TEXT_MAIN,
                              selectbackground=PRIMARY, selectforeground="#FFFFFF",
                              relief="flat", bd=0, font=("Segoe UI", 9),
                              activestyle="none", highlightthickness=0)
    file_list_sb = tk.Scrollbar(lst_box_frame, orient="vertical", command=file_listbox.yview)
    file_listbox.configure(yscrollcommand=file_list_sb.set)
    file_list_sb.pack(side="right", fill="y")
    file_listbox.pack(side="left", fill="both", expand=True, padx=4, pady=4)

    empty_hint = tk.Label(card_files, text="Click \"Add Files\" to select PDF or image tables",
                          font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_LIGHT)
    empty_hint.pack(anchor="w", pady=(0, 6))

    def update_file_ui():
        n = len(files)
        files_count_lbl.config(
            text=f"{n} file{'s' if n != 1 else ''} selected" if n else "0 files selected",
            bg="#EEF2FF" if n else "#F1F5F9",
            fg=PRIMARY if n else TEXT_MUTED
        )
        empty_hint.config(
            text=f"Total: {n} files queued for extraction" if n else
            "Click \"Add Files\" to select PDF or image tables"
        )

    def on_add():
        selected = filedialog.askopenfilenames(
            title="Select Tables (PDF, Images, Excel, CSV)",
            filetypes=[("All Supported", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp *.xlsx *.xls *.csv"),
                       ("Spreadsheets", "*.xlsx *.xls *.csv"),
                       ("PDF & Images", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
                       ("All files", "*.*")]
        )
        for f in selected:
            if f not in files:
                files.append(f)
                file_listbox.insert("end", f"  📄  {Path(f).name}")
        update_file_ui()

    def on_remove():
        for i in reversed(file_listbox.curselection()):
            files.pop(i)
            file_listbox.delete(i)
        update_file_ui()

    def on_clear():
        files.clear()
        file_listbox.delete(0, "end")
        update_file_ui()

    # Buttons Row
    files_btn_bar = tk.Frame(card_files, bg=CARD_BG)
    files_btn_bar.pack(fill="x")

    def make_button(parent, text, cmd, bg_col, fg_col, hover_col, pad_x=12):
        lbl = tk.Label(parent, text=text, font=("Segoe UI", 9, "bold"),
                       bg=bg_col, fg=fg_col, padx=pad_x, pady=5, cursor="hand2")
        lbl.bind("<Button-1>", lambda e: cmd())
        lbl.bind("<Enter>", lambda e: lbl.config(bg=hover_col))
        lbl.bind("<Leave>", lambda e: lbl.config(bg=bg_col))
        return lbl

    make_button(files_btn_bar, "＋ Add Files", on_add, PRIMARY, "#FFFFFF", PRIMARY_HOVER).pack(side="left", padx=(0, 6))
    make_button(files_btn_bar, "✕ Remove", on_remove, "#F1F5F9", "#475569", "#E2E8F0").pack(side="left", padx=(0, 6))
    make_button(files_btn_bar, "Clear all", on_clear, "#F1F5F9", "#475569", "#E2E8F0").pack(side="left")

    # ── CARD 2: OUTPUT & FORMAT ───────────────────────────────────────────────
    card_out = create_card(left_col)
    card_out.pack(fill="x", pady=(0, 10))

    tk.Label(card_out, text="Output Format & Destination", font=("Segoe UI", 10, "bold"),
             bg=CARD_BG, fg=TEXT_MAIN).pack(anchor="w", pady=(0, 8))

    # Segmented format pill selector
    fmt_frame = tk.Frame(card_out, bg=CARD_BG)
    fmt_frame.pack(fill="x", pady=(0, 8))
    tk.Label(fmt_frame, text="Export as:", font=("Segoe UI", 9), bg=CARD_BG, fg=TEXT_MUTED).pack(side="left", padx=(0, 10))

    pill_widgets = {}

    def set_format(val):
        fmt_var.set(val)
        for k, p in pill_widgets.items():
            if k == val:
                p.config(bg=PRIMARY, fg="#FFFFFF")
            else:
                p.config(bg="#F1F5F9", fg="#475569")

    for fkey, flabel in [("xlsx", "Excel (.xlsx)"), ("csv", "CSV (.csv)"), ("both", "Both")]:
        p_lbl = tk.Label(fmt_frame, text=f" {flabel} ", font=("Segoe UI", 9, "bold"),
                         padx=10, pady=4, cursor="hand2",
                         bg=PRIMARY if fkey == "xlsx" else "#F1F5F9",
                         fg="#FFFFFF" if fkey == "xlsx" else "#475569")
        p_lbl.bind("<Button-1>", lambda e, k=fkey: set_format(k))
        p_lbl.pack(side="left", padx=(0, 6))
        pill_widgets[fkey] = p_lbl

    # Output Folder
    dest_row = tk.Frame(card_out, bg=CARD_BG)
    dest_row.pack(fill="x", pady=(2, 0))
    tk.Label(dest_row, text="Folder:", font=("Segoe UI", 9), bg=CARD_BG, fg=TEXT_MUTED).pack(side="left", padx=(0, 8))

    dest_entry = tk.Entry(dest_row, textvariable=outdir, font=("Segoe UI", 9),
                          bg="#F8FAFC", fg=TEXT_MAIN, relief="flat",
                          highlightbackground=BORDER, highlightthickness=1)
    dest_entry.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=3)

    def on_browse_dest():
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            outdir.set(d)

    make_button(dest_row, "Browse…", on_browse_dest, "#F1F5F9", "#475569", "#E2E8F0", pad_x=10).pack(side="right")
    tk.Label(card_out, text="Leave blank to save directly next to each input file",
             font=("Segoe UI", 8), bg=CARD_BG, fg=TEXT_LIGHT).pack(anchor="w", pady=(2, 0))

    # ── CARD 3: EXTRACTION OPTIONS ────────────────────────────────────────────
    card_opts = create_card(left_col)
    card_opts.pack(fill="x", pady=(0, 12))

    tk.Label(card_opts, text="Extraction Options", font=("Segoe UI", 10, "bold"),
             bg=CARD_BG, fg=TEXT_MAIN).pack(anchor="w", pady=(0, 6))

    # Clean checkboxes in two neat rows
    chk_row1 = tk.Frame(card_opts, bg=CARD_BG)
    chk_row1.pack(fill="x", pady=(0, 4))
    ttk.Checkbutton(chk_row1, text="Numbers as real numbers", variable=numbers).pack(side="left", padx=(0, 16))
    ttk.Checkbutton(chk_row1, text="Merge into single sheet", variable=merge_var).pack(side="left")

    chk_row2 = tk.Frame(card_opts, bg=CARD_BG)
    chk_row2.pack(fill="x", pady=(0, 6))
    ttk.Checkbutton(chk_row2, text="★ CPT Mode (Clean Rates & Deduplicate)", variable=cpt_var).pack(side="left", padx=(0, 16))
    ttk.Checkbutton(chk_row2, text="Precise OCR", variable=precise).pack(side="left")

    chk_row3 = tk.Frame(card_opts, bg=CARD_BG)
    chk_row3.pack(fill="x", pady=(0, 6))
    ttk.Checkbutton(chk_row3, text="OPD rates only (hide ward/room/inpatient columns)", variable=opd_only).pack(side="left")

    # Advanced inline fields
    adv_frame = tk.Frame(card_opts, bg="#F8FAFC", highlightbackground=BORDER,
                         highlightthickness=1, padx=10, pady=8)
    adv_frame.pack(fill="x", pady=(4, 0))
    adv_frame.columnconfigure(1, weight=1)
    adv_frame.columnconfigure(3, weight=1)

    def add_adv_field(parent, label_text, var, row, col, width=8, hint=""):
        tk.Label(parent, text=label_text, font=("Segoe UI", 8, "bold"), bg="#F8FAFC",
                 fg=TEXT_MUTED).grid(row=row, column=col*2, sticky="w", padx=(0, 6), pady=2)
        box_f = tk.Frame(parent, bg="#F8FAFC")
        box_f.grid(row=row, column=col*2+1, sticky="w", pady=2)
        e = tk.Entry(box_f, textvariable=var, width=width, font=("Segoe UI", 9),
                     bg="#FFFFFF", fg=TEXT_MAIN, relief="flat",
                     highlightbackground=BORDER, highlightthickness=1)
        e.pack(side="left", ipady=2)
        if hint:
            tk.Label(box_f, text=f" {hint}", font=("Segoe UI", 8),
                     bg="#F8FAFC", fg=TEXT_LIGHT).pack(side="left")

    add_adv_field(adv_frame, "OCR Language:", lang_var, 0, 0, width=8, hint="e.g. eng, fil")
    add_adv_field(adv_frame, "Pages:", pages_var, 0, 1, width=8, hint="e.g. 1-3,7")
    add_adv_field(adv_frame, "Force Columns:", ncols_var, 1, 0, width=6, hint="(auto)")

    # OCR Mode selector
    tk.Label(adv_frame, text="OCR Mode:", font=("Segoe UI", 8, "bold"),
             bg="#F8FAFC", fg=TEXT_MUTED).grid(row=1, column=2, sticky="w", padx=(0, 6), pady=2)
    ocr_cb = ttk.Combobox(adv_frame, textvariable=ocr_var, values=["auto", "always", "never"],
                          width=7, state="readonly")
    ocr_cb.grid(row=1, column=3, sticky="w", pady=2)

    add_adv_field(adv_frame, "CPT Rate Col:", rate_var, 2, 0, width=8, hint="(blank = all rates)")
    add_adv_field(adv_frame, "Excel Sheet:", sheet_var, 2, 1, width=8, hint="(auto)")

    # ── BIG CONVERT BUTTON ────────────────────────────────────────────────────
    convert_btn_frame = tk.Frame(left_col, bg=BG)
    convert_btn_frame.pack(fill="x", pady=(0, 6))

    conv_btn = tk.Label(convert_btn_frame, text="⚡  Convert to Spreadsheet",
                        font=("Segoe UI", 11, "bold"), bg=PRIMARY, fg="#FFFFFF",
                        pady=11, cursor="hand2")
    conv_btn.pack(fill="x")

    def on_conv_enter(e):
        if not running.is_set():
            conv_btn.config(bg=PRIMARY_HOVER)

    def on_conv_leave(e):
        if not running.is_set():
            conv_btn.config(bg=PRIMARY)

    conv_btn.bind("<Enter>", on_conv_enter)
    conv_btn.bind("<Leave>", on_conv_leave)

    # ── RIGHT COLUMN: MODERN ACTIVITY TERMINAL ────────────────────────────────
    right_col = tk.Frame(root, bg=BG, padx=8, pady=4)
    right_col.grid(row=1, column=1, sticky="nsew")
    right_col.columnconfigure(0, weight=1)
    right_col.rowconfigure(1, weight=1)

    # Console Card
    console_card = tk.Frame(right_col, bg=CONSOLE_BG, highlightbackground="#334155",
                            highlightthickness=1, bd=0)
    console_card.grid(row=0, column=0, rowspan=2, sticky="nsew")
    console_card.columnconfigure(0, weight=1)
    console_card.rowconfigure(1, weight=1)

    # Console Title Bar (macOS / VS Code terminal style)
    c_bar = tk.Frame(console_card, bg=CONSOLE_BAR, padx=12, pady=8)
    c_bar.grid(row=0, column=0, sticky="ew")

    tk.Label(c_bar, text="●  ●  ●", font=("Segoe UI", 8), bg=CONSOLE_BAR,
             fg="#64748B").pack(side="left", padx=(0, 10))
    tk.Label(c_bar, text="ACTIVITY CONSOLE", font=("Segoe UI", 9, "bold"),
             bg=CONSOLE_BAR, fg="#94A3B8").pack(side="left")

    def on_clear_log():
        console_txt.configure(state="normal")
        console_txt.delete("1.0", "end")
        console_txt.configure(state="disabled")

    c_clear_lbl = tk.Label(c_bar, text="Clear", font=("Segoe UI", 8),
                           bg=CONSOLE_BAR, fg="#64748B", cursor="hand2")
    c_clear_lbl.bind("<Button-1>", lambda e: on_clear_log())
    c_clear_lbl.pack(side="right")

    # Terminal Text
    console_txt = tk.Text(console_card, bg=CONSOLE_BG, fg=CONSOLE_FG,
                          insertbackground=CONSOLE_FG, font=("Consolas", 9),
                          relief="flat", bd=10, wrap="word", state="disabled")
    c_sb = tk.Scrollbar(console_card, orient="vertical", command=console_txt.yview)
    console_txt.configure(yscrollcommand=c_sb.set)
    c_sb.grid(row=1, column=1, sticky="ns")
    console_txt.grid(row=1, column=0, sticky="nsew")

    console_txt.tag_configure("info", foreground=TAG_INFO)
    console_txt.tag_configure("ok",   foreground=TAG_OK)
    console_txt.tag_configure("warn", foreground=TAG_WARN)
    console_txt.tag_configure("err",  foreground=TAG_ERR)
    console_txt.tag_configure("muted",foreground=TEXT_LIGHT)

    # Welcome message in terminal
    console_txt.configure(state="normal")
    console_txt.insert("end", "table2sheet • v2.0 (Offline Mode)\n", "info")
    console_txt.insert("end", "Ready. Select files on the left and click Convert to begin.\n\n", "muted")
    console_txt.configure(state="disabled")

    # Progress bar and status
    prog_bar = ttk.Progressbar(right_col, mode="indeterminate", style="TProgressbar")
    prog_bar.grid(row=2, column=0, sticky="ew", pady=(8, 4))

    status_footer = tk.Frame(right_col, bg=BG)
    status_footer.grid(row=3, column=0, sticky="ew")

    status_label = tk.Label(status_footer, textvariable=status_msg, font=("Segoe UI", 9),
                            bg=BG, fg=TEXT_MUTED, anchor="w")
    status_label.pack(side="left")

    open_folder_btn = tk.Label(status_footer, text="📂 Open Output Folder", font=("Segoe UI", 9, "bold"),
                               bg="#EEF2FF", fg=PRIMARY, padx=8, pady=2, cursor="hand2")

    def on_open_folder(e):
        dest = last_dest.get()
        if dest and os.path.exists(dest):
            folder = dest if os.path.isdir(dest) else str(Path(dest).parent)
            os.startfile(folder)

    open_folder_btn.bind("<Button-1>", on_open_folder)

    # ── Logging Bridge ────────────────────────────────────────────────────────
    def log_say(msg: str):
        t_str = time.strftime("%H:%M:%S")
        tag = "ok" if "done" in msg.lower() or "saved" in msg.lower() else \
              "err" if "error" in msg.lower() else \
              "warn" if "warning" in msg.lower() else \
              "info" if msg.strip().startswith("Reading") else "muted"

        def _insert():
            console_txt.configure(state="normal")
            console_txt.insert("end", f"[{t_str}] ", "muted")
            console_txt.insert("end", f"{msg}\n", tag)
            console_txt.see("end")
            console_txt.configure(state="disabled")
            status_msg.set(msg[:70])
        root.after(0, _insert)

    # ── Background Worker ─────────────────────────────────────────────────────
    def run_conversion_thread():
        start_time = time.time()
        try:
            nc_str = ncols_var.get().strip()
            nc = int(nc_str) if nc_str.isdigit() else None
            o = Options(
                ocr=ocr_var.get(),
                numbers=numbers.get(),
                lang=lang_var.get().strip() or "eng",
                pages=pages_var.get().strip() or None,
                ncols=nc,
                precise=precise.get(),
                cpt=cpt_var.get(),
                opd_only=opd_only.get(),
                rate_col=rate_var.get().strip() or None,
                sheet=sheet_var.get().strip() or None,
                log=log_say,
            )
            saved = run(files, fmt_var.get(), outdir.get() or None, merge_var.get(), None, o)
            elapsed = time.time() - start_time
            if saved:
                last_dest.set(str(saved[0]))
                root.after(0, lambda: open_folder_btn.pack(side="right"))
            root.after(0, lambda: log_say(f"✔ Completed successfully in {elapsed:.1f}s"))
            root.after(0, lambda: status_msg.set(f"✔ Done in {elapsed:.1f}s"))
        except Exception as ex:  # noqa
            root.after(0, lambda: log_say(f"✘ ERROR: {ex}"))
            root.after(0, lambda: status_msg.set(f"Error: {ex}"))
        finally:
            running.clear()
            root.after(0, finish_run)

    def start_run():
        prog_bar.start(10)
        conv_btn.config(text="⏳  Converting… Please wait", bg="#64748B", cursor="wait")
        status_msg.set("Processing documents...")
        open_folder_btn.pack_forget()

    def finish_run():
        prog_bar.stop()
        conv_btn.config(text="⚡  Convert to Spreadsheet", bg=PRIMARY, cursor="hand2")

    def on_click_convert(e=None):
        if not files:
            messagebox.showinfo("table2sheet", "Please add at least one PDF or image file first.")
            return
        if running.is_set():
            return
        running.set()
        start_run()
        threading.Thread(target=run_conversion_thread, daemon=True).start()

    conv_btn.bind("<Button-1>", on_click_convert)

    root.mainloop()


if __name__ == "__main__":
    main()
