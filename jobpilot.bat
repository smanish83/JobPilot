@echo off
if not exist "%~dp0.venv\Scripts\python.exe" (
  echo JobPilot is not installed yet. Double-click setup_windows.bat first.
  exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0jobpilot.py" %*
