@echo off
title SAB Campus Excel Analyzer - Setup
echo ========================================
echo    SAB Campus Excel Analyzer
echo    First Time Setup
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed!
    echo.
    echo Please install Python 3.9+ from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: Check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
)

echo [1/3] Creating virtual environment...
python -m venv .venv

echo [2/3] Activating virtual environment...
call .venv\Scripts\activate.bat

echo [3/3] Installing dependencies (this may take a few minutes)...
pip install -r requirements.txt --quiet

echo.
echo ========================================
echo    Setup Complete!
echo ========================================
echo.
echo You can now run the app by double-clicking:
echo    Run_Excel_Analyzer.bat
echo.
pause
