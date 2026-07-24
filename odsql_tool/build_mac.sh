#!/bin/bash
# Builds OpenDentalSQLGenerator.app -- a standalone macOS app that needs
# nothing else installed to run. Run this ONCE (and again any time the
# source files change) on a Mac that has Python 3 installed (get it from
# python.org, or `brew install python` -- Tkinter is included either way).
#
# After it finishes, the finished program is at:
#   dist/OpenDentalSQLGenerator.app
#
# Note: since this isn't code-signed/notarized (that needs a paid Apple
# Developer account), macOS Gatekeeper will show an "unidentified
# developer" warning the first time someone opens it. Right-click ->
# Open (instead of double-clicking) gets past that once; after that it
# opens normally.
set -e
cd "$(dirname "$0")"

echo "Installing build dependencies (PyInstaller)..."
python3 -m pip install --upgrade pip
python3 -m pip install pyinstaller

echo
echo "Building OpenDentalSQLGenerator.app ..."
# permtype_details.json only exists once a PermType refresh has been run at
# least once here -- only bundle it if present, so a fresh checkout (which
# won't have it yet) still builds successfully.
EXTRA_DATA=()
if [ -f "permtype_details.json" ]; then
  EXTRA_DATA+=(--add-data "permtype_details.json:.")
fi
python3 -m PyInstaller --windowed --name OpenDentalSQLGenerator --distpath dist --workpath build --add-data "permtypes.json:." "${EXTRA_DATA[@]}" --add-data "prosperident_logo.png:." sql_generator_app.py

echo
echo "============================================================"
echo "Done. Your program is at:  dist/OpenDentalSQLGenerator.app"
echo "Copy that one file anywhere you like -- it runs standalone."
echo "============================================================"
