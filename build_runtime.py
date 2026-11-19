"""Build the portable Python runtime used by run.bat / launch.ps1.

Run this once on a Windows PC that has Python and the packages from
requirements.txt installed (system or user-level install):

    python build_runtime.py

It copies the running Python interpreter plus its installed packages into
./runtime/python next to this script. Zip the whole folder (including
runtime/) and share it; other laptops then need no Python install.
"""
import shutil
import site
import sys
from pathlib import Path

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

    print("3. Verifying portable runtime...")
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
