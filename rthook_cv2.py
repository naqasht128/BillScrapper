import sys
import types

# Ensure UTF-8 output streams in frozen environment
if sys.platform == "win32":
    try:
        if sys.stdout and hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if sys.stderr and hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# PyInstaller Runtime Hook: Stub cv2 so ddddocr does not pull in the 118MB opencv binary suite
if 'cv2' not in sys.modules:
    sys.modules['cv2'] = types.ModuleType('cv2')
