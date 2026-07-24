@echo off
REM ==========================================================================
REM  CRD Analysis - double-click launcher (Windows)
REM
REM  - Double-click this file to analyze Book1.xlsx.
REM  - Or drag any .xlsx file onto this file to analyze that file instead.
REM
REM  Results are written to the "analysis_output" folder next to this file.
REM  Requires Python 3 (https://www.python.org/downloads/ - tick
REM  "Add Python to PATH" during install).
REM ==========================================================================

REM Run from the folder this .bat lives in, whatever the current directory is.
cd /d "%~dp0"

REM Input file: the one dragged onto the .bat, or Book1.xlsx by default.
set "INPUT=%~1"
if "%INPUT%"=="" set "INPUT=Book1.xlsx"

if not exist "%INPUT%" (
    echo.
    echo Could not find the input file: "%INPUT%"
    echo Put your .xlsx in this folder, or drag it onto this .bat file.
    echo.
    pause
    exit /b 1
)

REM Prefer the Windows "py" launcher; fall back to "python".
where py >nul 2>nul
if %errorlevel%==0 (
    py analyze_crd.py "%INPUT%"
) else (
    python analyze_crd.py "%INPUT%"
)

if errorlevel 1 (
    echo.
    echo Something went wrong. If you saw "'python' is not recognized",
    echo install Python from https://www.python.org/downloads/ and tick
    echo "Add Python to PATH" during setup, then try again.
) else (
    echo.
    echo Done. Open the "analysis_output" folder for CRD_Analysis.xlsx and
    echo CRD_Analysis.html.
)
echo.
pause
