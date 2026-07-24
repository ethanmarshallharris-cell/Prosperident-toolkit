#!/bin/bash
# Builds DocumentToolkit.app -- a standalone macOS app that needs nothing
# else installed to run. Run this ONCE (and again any time the source
# files change) on a Mac that has Python 3 installed (get it from
# python.org, or `brew install python` -- Tkinter is included either way).
#
# After it finishes, the finished program is at: dist/DocumentToolkit.app
# That's what you hand out to staff -- no Python required on their
# machines, no installation, just double-click to run.
#
# Note: since this isn't code-signed/notarized (that needs a paid Apple
# Developer account), macOS Gatekeeper will show an "unidentified
# developer" warning the first time someone opens it. Right-click ->
# Open (instead of double-clicking) gets past that once; after that it
# opens normally.
set -e
cd "$(dirname "$0")"

echo "Installing build dependencies (PyMuPDF + PyInstaller)..."
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
python3 -m pip install pyinstaller

echo
echo "Building DocumentToolkit.app ..."
python3 -m PyInstaller --windowed --name DocumentToolkit --distpath dist --workpath build --add-data "src/prosperident_logo.png:." src/main_gui.py

echo
echo "============================================================"
echo "Done. Your program is at:  dist/DocumentToolkit.app"
echo "Copy that one file anywhere you like -- it runs standalone."
echo "============================================================"
