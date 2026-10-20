@echo off
REM Runs PDF to Excel straight from this folder (needs Python 3 with the
REM packages in requirements.txt). Staff normally use the built
REM PDFtoExcel.exe from the Tool Suite page instead.
cd /d "%~dp0"
python -m pip install --quiet -r requirements.txt
python pdf_to_excel_app.py %*
if errorlevel 1 (
    echo.
    echo Something went wrong. Make sure Python 3 is installed and on your PATH.
    echo You can download it from https://www.python.org/downloads/
    pause
)
