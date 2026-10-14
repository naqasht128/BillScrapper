#!/usr/bin/env bash
# Install Bill Scrapper to Linux desktop menu and application launcher
set -e

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
ICON_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/icons/hicolor/256x256/apps"

mkdir -p "$APP_DIR" "$ICON_DIR"

# Install application icon
cp "$PROJECT_DIR/assets/billscrapper-icon.png" "$ICON_DIR/billscrapper.png"

# Install .desktop launcher
DESKTOP_FILE="$APP_DIR/billscrapper.desktop"
cat <<EOF > "$DESKTOP_FILE"
[Desktop Entry]
Version=1.0
Type=Application
Name=Bill Scrapper
Comment=Automated PITC/PESCO Bill Scraper and MDI Audit Tool
Exec="$PROJECT_DIR/run.sh"
Path=$PROJECT_DIR
Icon=billscrapper
Terminal=false
Categories=Office;Utility;
StartupNotify=true
EOF

chmod +x "$DESKTOP_FILE"

# Update desktop database if tool is present
if command -v update-desktop-database &> /dev/null; then
    update-desktop-database "${XDG_DATA_HOME:-$HOME/.local/share}/applications" > /dev/null 2>&1 || true
fi

echo "Success! Bill Scrapper installed to Linux application menu ($DESKTOP_FILE)."
echo "You can now search and launch 'Bill Scrapper' from your desktop application launcher."
