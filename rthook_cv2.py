import sys
import types

# PyInstaller Runtime Hook: Stub cv2 so ddddocr does not pull in the 118MB opencv binary suite
if 'cv2' not in sys.modules:
    sys.modules['cv2'] = types.ModuleType('cv2')
