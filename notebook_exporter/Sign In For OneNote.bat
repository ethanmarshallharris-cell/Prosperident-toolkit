@echo off
REM First-time setup for the notebook exporter: signs this computer in to
REM OneNote so the exporter can read case notebooks. Run once per computer
REM (and again if the exporter later says the credential is invalid).
REM
REM SIGN IN WITH YOUR OWN PROSPERIDENT ACCOUNT. The on-screen text mentions a
REM "service account" - that wording belongs to the office automation; ignore
REM it. You will only be able to export notebooks you can already open.
REM Log: logs\Sign In For OneNote.log
setlocal
cd /d "%~dp0"
set "LOGFILE=%~dp0logs\%~n0.log"
if not exist "%~dp0logs" mkdir "%~dp0logs"
echo === %date% %time%  %~nx0 === >> "%LOGFILE%"
echo.
echo   Sign in with YOUR OWN Prosperident account (ignore any mention of a
echo   service account below). A browser code will be shown - follow it.
echo.
call "%~dp0_log.cmd" "%~dp0Get-DelegatedToken.ps1" -SignIn
echo.
echo   Log: %LOGFILE%
pause
exit /b 0
