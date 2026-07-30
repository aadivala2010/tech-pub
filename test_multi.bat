@echo off
setlocal
cd /d "%~dp0"
if "1"=="1" (
    echo inside
    pause
)
echo outside
pause
