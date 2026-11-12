@echo off
setlocal
title jobpilot setup
cd /d "%~dp0"
echo ============================================
echo   jobpilot setup for Windows
echo ============================================
echo.

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo Python was not found.
  echo Install Python 3.10 or newer from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during installation. Then run this file again.
  pause & exit /b 1
)
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 (
  echo Python 3.10 or newer is required. Please update Python from https://www.python.org/downloads/
  pause & exit /b 1
)

echo [1/3] Creating a private Python environment in this folder (.venv)...
if not exist ".venv\Scripts\python.exe" %PY% -m venv .venv
if errorlevel 1 ( echo Could not create the environment. & pause & exit /b 1 )

echo [2/3] Installing the libraries JobPilot needs (about a minute)...
".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet
if errorlevel 1 ( echo Installing the libraries failed. Check your internet connection and run this file again. & pause & exit /b 1 )
".venv\Scripts\python.exe" -m pip install pywin32 --quiet >nul 2>nul

echo [3/3] Guided setup
echo.
".venv\Scripts\python.exe" jobpilot.py init
echo.
echo Setup finished. Open a terminal in this folder and run, for example:
echo    jobpilot tailor "https://www.linkedin.com/jobs/view/1234567890/"
echo (In PowerShell type .\jobpilot instead of jobpilot.)
pause
