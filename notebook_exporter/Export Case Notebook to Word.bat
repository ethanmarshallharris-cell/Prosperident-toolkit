@echo off
REM Exports a whole case notebook to ONE Word document - cover page, contents,
REM every section and page, pictures embedded, attachments saved in a folder
REM beside it. Output: Documents\Prosperident Notebook Exports\
REM Reads only - nothing in the notebook is changed. Needs Word installed.
REM THE DOCUMENT HOLDS CLIENT AND PATIENT INFORMATION - file it in the case
REM document library; do not email it.
REM Log: logs\Export Case Notebook to Word.log
setlocal
cd /d "%~dp0"
set "LOGFILE=%~dp0logs\%~n0.log"
if not exist "%~dp0logs" mkdir "%~dp0logs"
echo === %date% %time%  %~nx0 === >> "%LOGFILE%"
call "%~dp0_log.cmd" "%~dp0Export-NotebookToWord.ps1" -ListCases
set "CASE="
set /p CASE=  Case name to export:
if "%CASE%"=="" (echo   No case entered - nothing exported. & pause & exit /b 1)
set "PDF="
set /p PDF=  Also make a PDF? (Y/N, blank = N):
if /i "%PDF%"=="Y" (
  call "%~dp0_log.cmd" "%~dp0Export-NotebookToWord.ps1" -CaseName "%CASE%" -Pdf
) else (
  call "%~dp0_log.cmd" "%~dp0Export-NotebookToWord.ps1" -CaseName "%CASE%"
)
echo.
echo   Log: %LOGFILE%
pause
exit /b 0
