@echo off
setlocal
cd /d "%~dp0"

REM -- 1. Python present? --
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

REM -- 2. Dependencies present? --
python -c "import flask, docx, fitz, requests" >nul 2>nul
if errorlevel 1 (
    echo Installing dependencies ^(first run only^)...
    python -m pip install --disable-pip-version-check -q -r requirements.txt
    if errorlevel 1 (
        echo.
        echo   Could not install the dependencies. Check your internet
        echo   connection and try running this again.
        echo.
        pause
        exit /b 1
    )
)

REM -- 3. Gemini API key present? --
if not exist ".env" goto askkey
findstr /b /c:"GEMINI_API_KEY=" ".env" >nul 2>nul
if not errorlevel 1 goto run

:askkey
echo.
echo   A Gemini API key is needed to run the agent (one time on this computer).
set /p GEMKEY="  Paste your Gemini API key and press Enter: "
if "%GEMKEY%"=="" (
    echo   No key entered - cannot continue.
    pause
    exit /b 1
)
> ".env" echo GEMINI_API_KEY=%GEMKEY%
echo   Saved. You won't be asked again on this computer.
echo.

REM -- 4. Launch --
:run
echo Starting DocRevise...
python server.py
pause
