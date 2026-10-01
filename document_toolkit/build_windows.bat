@echo off
REM Builds DocumentRedactor.exe -- a standalone Windows executable that
REM needs nothing else installed to run. Run this ONCE (and again any time
REM the source files change) on a Windows PC that has Python installed
REM (get it from python.org if needed -- any recent Python 3 works, and the
REM standard installer already includes tkinter, which this app needs).
REM
REM After it finishes, the finished program is at: dist\DocumentRedactor.exe
REM That single file is what you hand out to staff -- no Python required on
REM their machines, no installation, just double-click to run.

setlocal

REM Start from a clean slate every time: if a previous dist\DocumentRedactor.exe
REM or build\ folder is still sitting here, PyInstaller has to overwrite it
REM in place -- and on some Windows machines (antivirus real-time scanning a
REM freshly-written .exe, or the previous DocumentRedactor.exe still open
REM somewhere) that overwrite specifically is what triggers
REM "PermissionError: [WinError 5] Access is denied", even though a plain
REM fresh write to a brand-new file works fine. Deleting the old files
REM first, so PyInstaller is always creating rather than overwriting,
REM avoids that failure mode outright. If DocumentRedactor.exe is currently
REM running, close it before running this script -- a running program's
REM own .exe file can't be deleted out from under it on Windows.
if exist dist\DocumentRedactor.exe (
    echo Removing previous build output...
    del /f /q dist\DocumentRedactor.exe
)
if exist dist\DocumentRedactor.exe (
    echo.
    echo Could not remove the previous dist\DocumentRedactor.exe -- it's
    echo probably still running. Close it via Task Manager if it doesn't
    echo have a visible window, then run this script again.
    goto :error
)
if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__

REM Skipping "pip install --upgrade pip" here on purpose -- it adds a
REM network round-trip and a pip self-reinstall on every single build for
REM no real benefit to us; the pip that ships with a recent Python 3
REM installer is already more than sufficient for a plain "pip install -r
REM requirements.txt". If pip itself is genuinely too old on this machine,
REM run "python -m pip install --upgrade pip" once by hand and re-run this
REM script.
echo Installing build dependencies (PyMuPDF + openpyxl + PyInstaller)...
python -m pip install -r requirements.txt
python -m pip install pyinstaller
if errorlevel 1 goto :error

echo.
echo Building DocumentRedactor.exe ...
REM The --exclude-module flags below skip bundling optional libraries that
REM openpyxl/PyMuPDF can detect and use if present but don't actually need
REM for anything this app does (lxml, PIL, pandas, numpy, and lxml's own
REM optional HTML-parsing extras) -- if your Python environment happens to
REM have any of these installed for unrelated reasons, PyInstaller would
REM otherwise bundle them in "just in case", which measurably slows the
REM build and bloats the .exe for no benefit here.
python -m PyInstaller --onefile --windowed --name DocumentRedactor --distpath dist --workpath build --specpath build ^
  --exclude-module lxml --exclude-module PIL --exclude-module pandas ^
  --exclude-module numpy --exclude-module bs4 --exclude-module html5lib ^
  src\main_gui.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo Done. Your program is at:  dist\DocumentRedactor.exe
echo Copy that one file anywhere you like -- it runs standalone.
echo ============================================================
pause
goto :eof

:error
echo.
echo Something went wrong during the build -- see the messages above.
pause
exit /b 1
