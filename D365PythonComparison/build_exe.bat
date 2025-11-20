@echo off
REM Build standalone executable for D365 Comparison Tool
REM This creates a single .exe file that includes Python runtime

echo ======================================================================
echo   D365 Comparison Tool - EXE Builder
echo ======================================================================
echo.

REM Check Python installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    echo Please install Python first.
    pause
    exit /b 1
)

REM Check if PyInstaller is installed
echo Checking for PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo Failed to install PyInstaller!
        pause
        exit /b 1
    )
)
echo OK PyInstaller is ready

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist *.spec del /q *.spec

REM Build the executable
echo.
echo Building executable...
echo This may take a few minutes...
echo.

pyinstaller --onefile --console --name "D365ComparisonTool" main.py

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo   Build Successful!
    echo ======================================================================
    echo.
    echo Executable created at: dist\D365ComparisonTool.exe
    echo.
    echo You can now:
    echo   1. Run: dist\D365ComparisonTool.exe
    echo   2. Distribute the .exe file to other computers
    echo   3. No Python installation needed on target machines
    echo.
    echo Note: Some antivirus software may flag PyInstaller executables.
    echo       This is a false positive. You may need to add an exception.
    echo.
) else (
    echo.
    echo ======================================================================
    echo   Build Failed!
    echo ======================================================================
    echo.
    echo Please check the error messages above.
    echo.
)

pause
