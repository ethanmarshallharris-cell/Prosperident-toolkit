@echo off
REM _log.cmd - runs a PowerShell script and writes everything it prints to a
REM log file as well as the screen, so nothing has to be copied by hand.
REM
REM   usage (from a .bat that has set LOGFILE):
REM     call "%~dp0_log.cmd" "%~dp0Some-Script.ps1" -CaseName "Root Canals Rock Endodontics" -WhatIf
REM
REM   Arguments beginning with "-" are passed as switches; everything else is
REM   passed as a string. The script's exit code comes back as errorlevel.
setlocal
if not defined LOGFILE set "LOGFILE=%~dp0logs\unnamed.log"
for %%D in ("%LOGFILE%") do if not exist "%%~dpD" mkdir "%%~dpD"

set "PS1=%~1"
shift
set "ARGS="
:next
if "%~1"=="" goto run
set "A=%~1"
if "%A:~0,1%"=="-" goto switch
set "ARGS=%ARGS% '%A%'"
goto shifted
:switch
set "ARGS=%ARGS% %A%"
:shifted
shift
goto next

:run
for %%F in ("%PS1%") do set "NAME=%%~nxF"
echo.>>"%LOGFILE%"
echo --- %date% %time%  %NAME%  %ARGS% --->>"%LOGFILE%"
REM Tee-Object in Windows PowerShell 5.1 writes UTF-16, which cannot be mixed
REM with the batch file's own lines - so write UTF-8 by hand, line by line.
REM A script that dies on an unhandled error is caught here too, so the error
REM text lands in the log and the launcher sees a non-zero errorlevel.
REM AutoFlush (30 Sep 2026): each line reaches the log as it is printed, so a
REM long run can be followed - and a stuck one diagnosed - while it is going.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$w = New-Object System.IO.StreamWriter('%LOGFILE%', $true, (New-Object System.Text.UTF8Encoding $false)); $w.AutoFlush = $true; $rc = 0; try { & '%PS1%'%ARGS% *>&1 | ForEach-Object { $_; $w.WriteLine([string]$_) }; $rc = $LASTEXITCODE } catch { $m = 'SCRIPT STOPPED: ' + $_.Exception.Message + ' (' + $_.InvocationInfo.PositionMessage + ')'; Write-Host $m -ForegroundColor Red; $w.WriteLine($m); $rc = 1 } finally { $w.Close() }; exit $rc"
endlocal & exit /b %errorlevel%
