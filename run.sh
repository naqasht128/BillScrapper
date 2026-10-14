#!/usr/bin/env bash
# Bill Scrapper Linux Runner
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if ! command -v python3 &> /dev/null; then
    echo "Error: python3 is not installed. Please install Python 3 (e.g., sudo apt install python3 python3-tk python3-venv)."
    exit 1
fi

VENV_DIR="$SCRIPT_DIR/.venv"

# If virtual environment doesn't exist, create it and install requirements
if [ ! -d "$VENV_DIR" ]; then
    echo "First-time setup: Creating virtual environment in .venv..."
    python3 -m venv "$VENV_DIR"
    echo "Installing required Python packages..."
    "$VENV_DIR/bin/pip" install --upgrade pip
    "$VENV_DIR/bin/pip" install -r "$SCRIPT_DIR/requirements.txt"
fi

# Launch the application using the virtual environment
exec "$VENV_DIR/bin/python" "$SCRIPT_DIR/main.py" "$@"
