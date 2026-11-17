#!/usr/bin/env python3
"""Bill Scrapper - Application Launcher.
Bootstraps the application by delegating directly to the compiled native binary module (main.pyd).
"""

import sys
import types

# Lightweight cv2 stub to prevent ddddocr 1.6.1 from importing 118MB OpenCV
sys.modules.setdefault("cv2", types.ModuleType("cv2"))

import main

if __name__ == "__main__":
    main.start_app()
