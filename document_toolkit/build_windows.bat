@echo off
REM Builds DocumentToolkit.exe -- a standalone Windows executable that
REM needs nothing else installed to run. Run this ONCE (and again any time
REM the source files change) on a Windows PC that has Python installed
REM (get it from python.org if needed -- any recent Python 3 works, and the
REM standard installer already includes tkinter, which this app needs).
REM
REM After it finishes, the finished program is at: dist\DocumentToolkit.exe
REM That single file is what you hand out to staff -- no Python required on
REM their machines, no installation, just double-click to run.

setlocal

echo Installing build dependencies (PyMuPDF + PyInstaller)...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if errorlevel 1 goto :error

echo.
echo Building DocumentToolkit.exe ...
python -m PyInstaller --onefile --windowed --name DocumentToolkit --distpath dist --workpath build --specpath build --add-data "src\prosperident_logo.png;." src\main_gui.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo Done. Your program is at:  dist\DocumentToolkit.exe
echo Copy that one file anywhere you like -- it runs standalone.
echo ============================================================
pause
goto :eof

:error
echo.
echo Something went wrong during the build -- see the messages above.
pause
exit /b 1
