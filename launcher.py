#!/usr/bin/env python3
"""Bill Scrapper - Application Launcher.
Bootstraps the application by delegating directly to the compiled native binary module (main.pyd).
"""

import sys
import main

if __name__ == "__main__":
    main.start_app()
