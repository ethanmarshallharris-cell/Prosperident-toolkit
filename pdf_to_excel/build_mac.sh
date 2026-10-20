#!/bin/bash
# Builds PDFtoExcel.app -- a standalone macOS app. Run on a Mac with
# Python 3 installed. (GitHub Actions does this automatically on every
# release.) Result: dist/PDFtoExcel.app
# Not code-signed: the first time, right-click -> Open to get past the
# "unidentified developer" warning.
set -e
cd "$(dirname "$0")"
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller
python3 -m PyInstaller --windowed --name PDFtoExcel --distpath dist --workpath build --add-data "prosperident_logo.png:." pdf_to_excel_app.py
echo "Done. Your program is at: $(pwd)/dist/PDFtoExcel.app"
