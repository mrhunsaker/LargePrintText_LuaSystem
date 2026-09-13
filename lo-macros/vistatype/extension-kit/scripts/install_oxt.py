#!/usr/bin/env python3
from pathlib import Path
import shutil, subprocess
ROOT=Path(__file__).resolve().parents[1]
OXT=ROOT/"dist"/"vistatype-lp.oxt"
def find_unopkg():
    candidates=[shutil.which("unopkg"),"/usr/bin/unopkg","/usr/lib/libreoffice/program/unopkg","/Applications/LibreOffice.app/Contents/MacOS/unopkg",r"C:\Program Files\LibreOffice\program\unopkg.exe",r"C:\Program Files (x86)\LibreOffice\program\unopkg.exe"]
    for p in candidates:
        if p and Path(p).exists(): return str(p)
    raise SystemExit("Could not locate unopkg; add LibreOffice's program directory to PATH.")
def main():
    if not OXT.exists(): raise SystemExit("Run build_oxt.py first")
    subprocess.run([find_unopkg(),"add",str(OXT)],check=True)
    print("Installed. Fully restart LibreOffice before testing.")
if __name__=="__main__": main()
