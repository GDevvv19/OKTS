#!/bin/bash
# Run this ONCE on the Mac where you want the app (e.g. the Apple Silicon / M3
# Mac). It builds a native, double-clickable ShipReportBuilder.app for THIS Mac's
# processor. No changes needed on any other machine.
cd "$(dirname "$0")" || exit 1

echo "Building a native ShipReportBuilder.app for this Mac..."
echo "(first time downloads the build tools; needs internet)"
echo ""

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is required. Run 'xcode-select --install' or install from python.org, then retry."
  read -r -p "Press Enter to close."; exit 1
fi

python3 -m venv .build-venv || { echo "Could not create build environment."; read -r -p "Press Enter to close."; exit 1; }
./.build-venv/bin/pip install --quiet --upgrade pip pyinstaller openpyxl tzdata || {
  echo "Could not install build tools (need internet)."; read -r -p "Press Enter to close."; exit 1; }

ICON_ARG=""
[ -f appicon.icns ] && ICON_ARG="--icon appicon.icns"
./.build-venv/bin/pyinstaller --noconfirm --clean --windowed \
  --name "ShipReportBuilder" --paths src --collect-submodules openpyxl --collect-data tzdata $ICON_ARG \
  src/webapp.py || { echo "Build failed."; read -r -p "Press Enter to close."; exit 1; }

echo ""
echo "============================================================"
echo "Done!  Your native app is here:"
echo "   $(pwd)/dist/ShipReportBuilder.app"
echo "Drag it into your Applications folder."
echo "First open: right-click the app -> Open -> Open."
echo "============================================================"
open dist 2>/dev/null
read -r -p "Press Enter to close."
