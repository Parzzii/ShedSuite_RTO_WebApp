@echo off
setlocal
cd /d "%~dp0"

echo.
echo ================================================
echo   ShedSuite - RTO Pro Web App V6.3.6
echo   V3 base + RTO Pro upload + Delivery Certificate
echo ================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  set "PY=python"
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating local Python environment...
  %PY% -m venv .venv
  if errorlevel 1 goto :error
)

echo Installing/updating required packages...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :error

if not exist ".playwright_ready" (
  echo Installing Playwright Chromium for Delivery Certificate only...
  ".venv\Scripts\python.exe" -m playwright install chromium
  if errorlevel 1 goto :error
  echo ready> .playwright_ready
)

if not exist ".env" copy /y ".env.example" ".env" >nul

echo.
echo Starting local web app at http://127.0.0.1:5050
echo Keep this window open while you use the app.
start "" /b ".venv\Scripts\python.exe" -c "import time,webbrowser; time.sleep(2); webbrowser.open('http://127.0.0.1:5050')"
".venv\Scripts\python.exe" app.py
goto :end

:error
echo.
echo Setup failed. Make sure Python 3 is installed and this PC has internet access for the first package install.
pause

:end
endlocal
