@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo   Python is not installed, or not on your PATH.
    echo   Install Python 3 from https://www.python.org/downloads/
    echo   ^(tick "Add python.exe to PATH" during setup^), then run this again.
    echo.
    pause
    exit /b 1
)

python -c "import flask, docx, fitz, requests" >nul 2>nul
if errorlevel 1 (
    echo Installing dependencies (first run only)...
    python -m pip install --disable-pip-version-check -q -r requirements.txt
    if errorlevel 1 (
        echo.
        echo   Could not install the dependencies.
        echo.
        pause
        exit /b 1
    )
)

if not exist ".env" goto askkey
findstr /b /c:"GEMINI_API_KEY=" ".env" >nul 2>nul
if not errorlevel 1 goto run

:askkey
echo.
echo   A Gemini API key is needed.
set /p GEMKEY="  Paste your Gemini API key: "
if "%GEMKEY%"=="" (
    echo   No key entered.
    pause
    exit /b 1
)
> ".env" echo GEMINI_API_KEY=%GEMKEY%
echo   Saved.
echo.

:run
echo Starting DocRevise...
python server.py
pause
