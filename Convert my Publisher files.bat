@echo off
cd /d "%~dp0"

REM === EDIT THESE TWO FOLDERS ===
REM Use "." for the folder this .bat file lives in.
REM Or paste full paths, for example: C:\Users\Jane\OneDrive\Documents\MyPubs
set "SOURCE=."
set "OUTPUT=."

REM Optional: seconds to wait for export (600 = 10 minutes).
REM Leave commented out to use the script default (180).
REM set "TIMEOUT=600"

if defined TIMEOUT (
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "%SOURCE%" -OutputRoot "%OUTPUT%" -Skip -ExportTimeoutSeconds %TIMEOUT%
) else (
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0pub2pdf.ps1" -SourceRoot "%SOURCE%" -OutputRoot "%OUTPUT%" -Skip
)

echo.
echo Finished. Press any key to close this window.
pause >nul
