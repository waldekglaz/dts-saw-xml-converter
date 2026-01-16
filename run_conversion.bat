@echo off
echo ========================================================
echo      DTS-SAW Excel to XML Converter
echo ========================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not found in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo IMPORANT: Check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b
)

echo Installing necessary libraries (this only happens once)...
pip install -r requirements.txt >nul 2>&1
if %errorlevel% neq 0 (
    echo Error installing requirements. Please check your internet connection.
    pause
    exit /b
)

echo.
echo Starting Conversion...
echo.

python dts-saw.py

echo.
echo ========================================================
echo Process Complete.
echo You can close this window now.
pause
