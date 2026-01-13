@echo off
title SAB Campus Excel Analyzer
cd /d "%~dp0"

REM Check if .venv exists
if not exist ".venv" (
    echo Virtual environment not found!
    echo Please run SETUP_FIRST.bat first.
    pause
    exit /b 1
)

REM Activate virtual environment
call .venv\Scripts\activate.bat

echo ========================================
echo    SAB Campus Excel Analyzer
echo ========================================
echo.
echo Starting application...
echo Browser will open automatically.
echo.
echo Press Ctrl+C to stop the server.
echo ========================================

streamlit run app.py --server.headless false --browser.gatherUsageStats false
