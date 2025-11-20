@echo off
REM Setup script for D365 Python Comparison Tool
REM Run this to install dependencies

echo ======================================================================
echo   D365 Python Comparison Tool - Setup
echo ======================================================================
echo.

REM Check Python installation
echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   X Python not found! Please install Python 3.8 or higher.
    echo     Download from: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo   OK Found: %PYTHON_VERSION%

REM Check pip installation
echo Checking pip installation...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   X pip not found!
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('pip --version 2^>^&1') do set PIP_VERSION=%%i
echo   OK Found pip

echo.
echo Installing required Python packages...
echo This may take a few moments...
echo.

REM Install requirements
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo   X Failed to install packages!
    echo.
    pause
    exit /b 1
)

echo.
echo ======================================================================
echo   Setup Complete!
echo ======================================================================
echo.
echo To run the tool, execute:
echo   run.bat
echo.
echo OR double-click run.bat
echo.
echo For more information, see README.md
echo.
pause
