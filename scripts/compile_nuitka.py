#!/usr/bin/env python3
"""Bill Scrapper - Nuitka Binary Compiler Script.
Compiles main.py into a native C++ (.pyd) binary.
Protects source code against decompilers, bytecode extractors, and AI models.
"""

import os
import sys
import glob
import subprocess
from pathlib import Path

# Ensure UTF-8 stdout/stderr on Windows to prevent UnicodeEncodeError
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT_DIR = Path(__file__).resolve().parent.parent
MAIN_PY = ROOT_DIR / "main.py"


def compile_with_nuitka():
    if os.environ.get("SKIP_NUITKA") == "1":
        print("[INFO] SKIP_NUITKA=1 detected. Skipping Nuitka compilation and using standard source files.")
        return

    if not MAIN_PY.exists():
        print("[WARN] main.py does not exist. Already compiled?")
        return

    print("=" * 60)
    print("[SECURE] NUITKA COMPILER: Securing Bill Scrapper core logic to native C++ (.pyd)")
    print("=" * 60)

    cmd = [
        sys.executable, "-m", "nuitka",
        "--module",
        "--assume-yes-for-downloads",
        f"--output-dir={str(ROOT_DIR)}",
        "--remove-output",
        "--no-progressbar",
        str(MAIN_PY)
    ]

    print("  -> Compiling main.py to native C++ machine code...")
    result = subprocess.run(cmd, cwd=str(ROOT_DIR))
    if result.returncode != 0:
        print(f"[ERROR] Failed to compile main.py with Nuitka (exit code {result.returncode})")
        sys.exit(result.returncode)

    # Find the generated .pyd file
    pyd_matches = list(ROOT_DIR.glob("main*.pyd"))
    if not pyd_matches:
        print("[ERROR] Compiled .pyd for main not found in project directory.")
        sys.exit(1)

    print(f"  [OK] Generated native C++ binary: {pyd_matches[0].name}")

    # Remove original main.py so only .pyd is bundled by PyInstaller
    MAIN_PY.unlink()
    print("  [OK] Removed original source file: main.py")

    print("\n" + "=" * 60)
    print("[DONE] NUITKA COMPILATION COMPLETED: Core application logic is 100% secured!")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    compile_with_nuitka()
