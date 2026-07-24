@echo off
cd /d "%~dp0"
python sql_generator_app.py
if errorlevel 1 (
    echo.
    echo Something went wrong. Make sure Python 3 is installed and on your PATH.
    echo You can download it from https://www.python.org/downloads/
    pause
)
