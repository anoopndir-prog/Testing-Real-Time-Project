#!/usr/bin/env bash
# Build SKF Report Generator as a macOS .app bundle using PyInstaller.
# Run from the project root:  bash build_mac_app.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV="$SCRIPT_DIR/.venv"
PYTHON="$VENV/bin/python"
PIP="$VENV/bin/pip"
PYINSTALLER="$VENV/bin/pyinstaller"

echo "=== SKF Report Generator – macOS Build ==="
echo "Python: $($PYTHON --version)"

# Install / upgrade build-time deps
echo ""
echo "--- Checking dependencies ---"
"$PIP" install --quiet --upgrade pyinstaller

# Optional: install drag-and-drop support if available
"$PIP" install --quiet tkinterdnd2 2>/dev/null && echo "tkinterdnd2 installed (drag-and-drop enabled)" || echo "tkinterdnd2 unavailable – drag-and-drop disabled"

# Optional: install calendar widget
"$PIP" install --quiet tkcalendar 2>/dev/null && echo "tkcalendar installed (calendar picker enabled)" || echo "tkcalendar unavailable – using built-in calendar"

# Clean previous build artefacts
echo ""
echo "--- Cleaning previous build ---"
rm -rf build dist

# Run PyInstaller
echo ""
echo "--- Building .app bundle ---"
"$PYINSTALLER" --clean --noconfirm SKFReportGenerator.spec

APP_PATH="$SCRIPT_DIR/dist/SKF Report Generator.app"

if [ -d "$APP_PATH" ]; then
    echo ""
    echo "=== Build successful ==="
    echo "App bundle: $APP_PATH"
    echo ""
    echo "To install: drag '$APP_PATH' to your Applications folder."
    echo "To run now: open \"$APP_PATH\""
else
    echo ""
    echo "ERROR: Build failed – .app not found in dist/"
    exit 1
fi
