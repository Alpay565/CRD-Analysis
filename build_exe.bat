@echo off
REM ==========================================================================
REM  Build CRD_Analyzer.exe  (run this once, on Windows)
REM
REM  Produces a single standalone application file:  dist\CRD_Analyzer.exe
REM  You can then double-click that .exe - no Python needed on the machine
REM  that runs it.
REM
REM  Requires Python 3 with "Add Python to PATH" ticked:
REM    https://www.python.org/downloads/
REM ==========================================================================

cd /d "%~dp0"

echo.
echo === Checking Python ===
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")
%PY% --version
if errorlevel 1 (
    echo.
    echo Python was not found. Install it from https://www.python.org/downloads/
    echo and tick "Add Python to PATH" during setup, then run this again.
    pause
    exit /b 1
)

echo.
echo === Installing PyInstaller (one-time) ===
%PY% -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo.
    echo Could not install PyInstaller. If you are behind a corporate proxy,
    echo ask IT for pip access, or use the GitHub Actions build instead
    echo ^(see README: "Getting the app"^).
    pause
    exit /b 1
)

echo.
echo === Building CRD_Analyzer.exe ===
%PY% -m PyInstaller --onefile --noconsole --name CRD_Analyzer ^
    --hidden-import analyze_crd crd_app.py
if errorlevel 1 (
    echo.
    echo Build failed. See the messages above.
    pause
    exit /b 1
)

echo.
echo ==========================================================
echo  Done.  Your application is here:
echo     %~dp0dist\CRD_Analyzer.exe
echo.
echo  Double-click it to start the CRD Analyzer.
echo ==========================================================
echo.
pause
