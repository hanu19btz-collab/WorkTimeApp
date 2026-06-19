@echo off
title Work Time Report
cd /d "%~dp0"
echo.
echo  =========================================
echo   Work Time Report  v2.0
echo  =========================================
echo.

where streamlit >nul 2>&1
if errorlevel 1 (
    echo  [!] Streamlit not found. Installing dependencies...
    echo.
    pip install -r requirements.txt
    echo.
)

echo  Starting app...  Press Ctrl+C to stop.
echo.
streamlit run app.py --server.headless false
pause
