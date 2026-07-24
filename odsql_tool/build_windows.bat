@echo off
REM Builds OpenDentalSQLGenerator.exe -- a standalone Windows executable
REM that needs nothing else installed to run. Run this ONCE (and again
REM any time the source files change) on a Windows PC that has Python
REM installed (get it from python.org if needed -- any recent Python 3
REM works, and the standard installer already includes tkinter, which
REM this app needs).
REM
REM After it finishes, the finished program is at:
REM   dist\OpenDentalSQLGenerator.exe
REM That single file is what you hand out to staff -- no Python required
REM on their machines, no installation, just double-click to run. Its
REM current permtypes.json / permtype_details.json are bundled in and
REM copied out next to the .exe the first time it runs, then kept current
REM from there exactly like the plain-script version already does.

setlocal

echo Installing build dependencies (PyInstaller)...
python -m pip install --upgrade pip
python -m pip install pyinstaller
if errorlevel 1 goto :error

echo.
echo Building OpenDentalSQLGenerator.exe ...
python -m PyInstaller --onefile --windowed --name OpenDentalSQLGenerator --distpath dist --workpath build --specpath build --add-data "permtypes.json;." --add-data "permtype_details.json;." --add-data "prosperident_logo.png;." sql_generator_app.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo Done. Your program is at:  dist\OpenDentalSQLGenerator.exe
echo Copy that one file anywhere you like -- it runs standalone.
echo ============================================================
pause
goto :eof

:error
echo.
echo Something went wrong during the build -- see the messages above.
pause
exit /b 1
