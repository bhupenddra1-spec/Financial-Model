@echo off
REM Double-click to start the Compliance Tools (MSME Verification + Struck-Off Companies).
REM To use your own data files or an API, set MSME_PROVIDER / MCA_PROVIDER here (see README.md).
cd /d "%~dp0"
echo Installing / checking required packages...
python -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo.
  echo Python was not found or the install failed. Install Python from https://www.python.org/downloads
  echo and tick "Add python.exe to PATH" during setup.
  pause
  exit /b 1
)
start "" http://127.0.0.1:5000/struck-off
echo.
echo The tool is running at http://127.0.0.1:5000  -  keep this window open. Press Ctrl+C to stop.
python -m msme_verifier.app
pause
