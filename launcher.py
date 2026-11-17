#!/usr/bin/env python3
"""Bill Scrapper - Application Launcher.
Bootstraps the application by delegating directly to the compiled native binary module (main.pyd).
"""

import sys
import types

# Ensure UTF-8 output streams on Windows
if sys.platform == "win32":
    try:
        if sys.stdout and hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if sys.stderr and hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Lightweight cv2 stub to prevent ddddocr 1.6.1 from importing 118MB OpenCV
sys.modules.setdefault("cv2", types.ModuleType("cv2"))

import main

if __name__ == "__main__":
    main.start_app()
