@echo off
REM Runs PDF to Excel straight from this folder (needs Python 3 with the
REM packages in requirements.txt). Staff normally use the built
REM PDFtoExcel.exe from the Tool Suite page instead.
cd /d "%~dp0"
python -m pip install --quiet -r requirements.txt
if not exist "tessdata\eng.traineddata" (
    echo Downloading the OCR language file ^(about 23 MB, one time only^)...
    if not exist tessdata mkdir tessdata
    powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/tesseract-ocr/tessdata/raw/4.1.0/eng.traineddata' -OutFile 'tessdata\eng.traineddata'"
)
python pdf_to_excel_app.py %*
if errorlevel 1 (
    echo.
    echo Something went wrong. Make sure Python 3 is installed and on your PATH.
    echo You can download it from https://www.python.org/downloads/
    pause
)
