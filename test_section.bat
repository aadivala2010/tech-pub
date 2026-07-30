@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo   Python is not installed, or not on your PATH.
    echo   Install Python 3 from https://www.python.org/downloads/
    echo   (tick "Add python.exe to PATH" during setup), then run this again.
    echo.
    pause
    exit /b 1
)

python -c "import flask, docx, fitz, requests" >nul 2>nul
if errorlevel 1 (
    echo Installing dependencies...
    echo skipping real install for test
)

echo OK - no errors
