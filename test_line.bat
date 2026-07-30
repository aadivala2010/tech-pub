@echo on
setlocal
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 ( echo FAIL ) else ( echo HAVE )
python -c "import flask" >nul 2>nul
if errorlevel 1 ( echo MISSING ) else ( echo HAVE )
if not exist ".env" goto askkey
findstr /b /c:"GEMINI_API_KEY=" ".env" >nul 2>nul
if not errorlevel 1 goto run
:askkey
echo ASKKEY
set /p GEMKEY="Enter key: "
if "%GEMKEY%"=="" ( echo NO KEY & pause & exit /b 1 )
> ".env" echo GEMINI_API_KEY=%GEMKEY%
echo SAVED
:run
echo RUNNING
