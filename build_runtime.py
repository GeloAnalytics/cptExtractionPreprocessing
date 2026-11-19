"""Build the portable Python runtime used by run.bat / launch.ps1.

Run this once on a Windows PC that has Python and the packages from
requirements.txt installed (system or user-level install):

    python build_runtime.py

It copies the running Python interpreter plus its installed packages into
./runtime/python, and an installed Tesseract OCR into ./runtime/tesseract
(where table2sheet looks for it). Zip the whole folder (including runtime/)
and share it; other laptops then need no Python or Tesseract install.
"""
import os
import shutil
import site
import sys
from pathlib import Path
from typing import Optional

APP_DIR = Path(__file__).resolve().parent
src_python = Path(sys.base_prefix)
dest_runtime = APP_DIR / "runtime" / "python"

SKIP_TOP_LEVEL = {"Doc", "include", "NEWS.txt"}


def ignore_patterns(path, names):
    ignored = {n for n in names if n == "__pycache__"}
    if Path(path) == src_python:
        ignored |= SKIP_TOP_LEVEL & set(names)
    return ignored


def copy_tree_contents(src: Path, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    for item in src.iterdir():
        if item.name in SKIP_TOP_LEVEL or item.name == "__pycache__":
            continue
        target = dest / item.name
        if target.exists():
            continue
        if item.is_dir():
            shutil.copytree(item, target, ignore=ignore_patterns)
        else:
            shutil.copy2(item, target)


def find_tesseract_dir() -> Optional[Path]:
    """Locate an installed Tesseract-OCR folder (the one holding tesseract.exe + tessdata)."""
    candidates = []
    if os.environ.get("TESSERACT_CMD"):
        candidates.append(Path(os.environ["TESSERACT_CMD"]).parent)
    found = shutil.which("tesseract")
    if found:
        candidates.append(Path(found).resolve().parent)
    for env in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        root = os.environ.get(env)
        if root:
            candidates += [Path(root) / "Tesseract-OCR", Path(root) / "Programs" / "Tesseract-OCR"]
    for c in candidates:
        if (c / "tesseract.exe").exists() and (c / "tessdata").is_dir():
            return c
    return None


def bundle_tesseract():
    dest = APP_DIR / "runtime" / "tesseract"
    if (dest / "tesseract.exe").exists():
        print(f"3. Tesseract already bundled at {dest}, skipping.")
        return
    src = find_tesseract_dir()
    if not src:
        print("3. WARNING: Tesseract OCR not found; scans/photos will not work from this runtime.")
        print("   Install it (winget install UB-Mannheim.TesseractOCR) and re-run this script.")
        return
    print(f"3. Copying Tesseract OCR from {src}...")
    shutil.copytree(src, dest, ignore=shutil.ignore_patterns("unins*", "*.html", "doc"))


def main():
    print(f"Source Python:     {src_python}")
    print(f"Portable runtime:  {dest_runtime}")

    print("1. Copying base Python files (includes system site-packages)...")
    copy_tree_contents(src_python, dest_runtime)

    dest_sp = dest_runtime / "Lib" / "site-packages"
    user_sp = Path(site.getusersitepackages())
    if user_sp.is_dir():
        print(f"2. Copying user site-packages from {user_sp}...")
        copy_tree_contents(user_sp, dest_sp)
    else:
        print("2. No user site-packages found, skipping.")

    bundle_tesseract()

    print("4. Verifying portable runtime...")
    portable_py = dest_runtime / "python.exe"
    print(f"   python.exe present: {portable_py.exists()}")
    missing = []
    for mod in ("pdfplumber", "openpyxl", "numpy", "cv2", "PIL", "pytesseract"):
        if not any(dest_sp.glob(f"{mod}*")):
            missing.append(mod)
    if missing:
        print(f"   WARNING: not found in runtime site-packages: {', '.join(missing)}")
        print("   Install them first:  python -m pip install --user -r requirements.txt")
    else:
        print("   All required packages present.")


if __name__ == "__main__":
    main()
