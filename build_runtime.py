import os, sys, shutil
from pathlib import Path

src_python = Path(r"C:\Users\angelo.urriza\AppData\Local\Programs\Python\Python314")
src_roaming = Path(r"C:\Users\angelo.urriza\AppData\Roaming\Python\Python314\site-packages")
dest_runtime = Path(r"c:\Users\angelo.urriza\Downloads\table2sheet_app\runtime\python")

print(f"Creating portable runtime at: {dest_runtime}")
dest_runtime.mkdir(parents=True, exist_ok=True)

# Ignore Doc, include, Scripts if unneeded to keep zip lightweight
def ignore_patterns(path, names):
    ignored = set()
    p = Path(path)
    if p == src_python:
        for n in ["Doc", "include", "NEWS.txt"]:
            if n in names:
                ignored.add(n)
    for n in names:
        if n == "__pycache__":
            ignored.add(n)
    return ignored

print("1. Copying base Python files...")
for item in src_python.iterdir():
    if item.name in ["Doc", "include", "NEWS.txt"]:
        continue
    target = dest_runtime / item.name
    if not target.exists():
        if item.is_dir():
            shutil.copytree(item, target, ignore=ignore_patterns)
        else:
            shutil.copy2(item, target)

print("2. Copying user site-packages (openpyxl, numpy, pandas)...")
dest_sp = dest_runtime / "Lib" / "site-packages"
dest_sp.mkdir(parents=True, exist_ok=True)

for item in src_roaming.iterdir():
    if item.name == "__pycache__":
        continue
    target = dest_sp / item.name
    if not target.exists():
        if item.is_dir():
            shutil.copytree(item, target, ignore=ignore_patterns)
        else:
            shutil.copy2(item, target)

print("3. Verifying portable runtime...")
portable_py = dest_runtime / "python.exe"
print(f"Portable python exists: {portable_py.exists()}")

