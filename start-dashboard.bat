@echo off
chcp 65001 >nul
title Infutive Payroll Dashboard
cd /d "%~dp0"

echo.
echo ============================================
echo   Infutive Payroll Dashboard - Starting...
echo ============================================
echo.

where py >nul 2>&1
if %errorlevel%==0 (
  set PY=py
) else (
  set PY=python
)

echo Using: %PY%
%PY% --version
if errorlevel 1 (
  echo.
  echo ERROR: Python not found. Reinstall from python.org
  echo          and tick "Add python.exe to PATH".
  pause
  exit /b 1
)

if not exist "payroll-dashboard\config\payroll.xml" (
  echo.
  echo ERROR: payroll.xml missing.
  echo Run from the folder that contains payroll_dashboard and payroll-dashboard.
  pause
  exit /b 1
)

echo.
echo Installing required packages (first time may take 1-2 minutes)...
%PY% -m pip install -q -r requirements-dashboard.txt
if errorlevel 1 (
  echo pip install failed. Try: %PY% -m pip install -r requirements-dashboard.txt
  pause
  exit /b 1
)

echo.
echo Starting web server... Do NOT close this window.
echo Open browser: http://localhost:8080
echo.

start "" "http://localhost:8080"
%PY% -m uvicorn payroll_dashboard.server:app --host 127.0.0.1 --port 8080

pause
