@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PYTHON=py -3") else (set "PYTHON=python")
if not exist .venv\Scripts\python.exe (
  echo Creating the local Python environment...
  %PYTHON% -m venv .venv
  if errorlevel 1 (echo Python 3.10+ is required. & pause & exit /b 1)
)
echo Installing/updating dependencies...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (echo Could not install the required Python packages. & pause & exit /b 1)
echo Starting local web app...
.venv\Scripts\python.exe app.py
pause
