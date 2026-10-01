@echo off
REM Double-click to release changes to the Document Toolkit and/or OpenDental SQL Generator.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Ship-Update.ps1"
echo.
pause
