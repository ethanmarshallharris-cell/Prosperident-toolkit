@echo off
REM Builds PDFtoExcel.exe -- a standalone Windows program that needs
REM nothing else installed. Run on a Windows PC with Python 3 installed.
REM (GitHub Actions does this automatically on every release; this is for
REM building by hand.)  Result: dist\PDFtoExcel.exe
setlocal
cd /d "%~dp0"
echo Installing build dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :error
if not exist "tessdata\eng.traineddata" (
    echo Downloading the OCR language file ^(about 23 MB, one time only^)...
    if not exist tessdata mkdir tessdata
    powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/tesseract-ocr/tessdata/raw/4.1.0/eng.traineddata' -OutFile 'tessdata\eng.traineddata'"
)
if not exist "tessdata\eng.traineddata" goto :error
echo.
echo Building PDFtoExcel.exe ...
python -m PyInstaller --onefile --windowed --name PDFtoExcel --distpath dist --workpath build --add-data "prosperident_logo.png;." --add-data "tessdata;tessdata" pdf_to_excel_app.py
if errorlevel 1 goto :error
echo.
echo ============================================================
echo Done. Your program is at:  %~dp0dist\PDFtoExcel.exe
echo ============================================================
pause
goto :eof
:error
echo.
echo Something went wrong during the build -- see the messages above.
pause
exit /b 1
