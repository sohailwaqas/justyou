@echo off
rem Double-click to start handy on Windows. First run sets everything up.
cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" goto start

echo First-time setup - this takes a few minutes, please wait...
py -3.11 -m venv .venv 2>nul
if exist ".venv\Scripts\activate.bat" goto install
py -3.12 -m venv .venv 2>nul
if exist ".venv\Scripts\activate.bat" goto install
python -m venv .venv 2>nul
if exist ".venv\Scripts\activate.bat" goto install
echo.
echo Python was not found. Install Python 3.11 from python.org
echo and tick "Add python.exe to PATH" in the installer, then run this again.
pause
exit /b 1

:install
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installation failed. Check your internet connection and try again.
  rmdir /s /q .venv
  pause
  exit /b 1
)
goto run

:start
call ".venv\Scripts\activate.bat"

:run
echo.
echo Starting handy - your browser will open at http://localhost:8501
echo Close this window (or press Ctrl+C) to stop the app.
streamlit run app.py
pause
