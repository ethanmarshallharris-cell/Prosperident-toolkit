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
python3 -m PyInstaller --windowed --name OpenDentalSQLGenerator --distpath dist --workpath build --specpath build --add-data "permtypes.json:." --add-data "permtype_details.json:." --add-data "prosperident_logo.png:." sql_generator_app.py

echo
echo "============================================================"
echo "Done. Your program is at:  dist/OpenDentalSQLGenerator.app"
echo "Copy that one file anywhere you like -- it runs standalone."
echo "============================================================"
