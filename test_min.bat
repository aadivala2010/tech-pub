@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed, or not on your PATH.
    pause
    exit /b 1
)

python server.py
pause
