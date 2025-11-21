# Building a Standalone .EXE File

If you want to create a standalone `.exe` file that doesn't require Python to be installed on the target machine, you can use **PyInstaller**.

## Method 1: Simple EXE (Recommended)

### Step 1: Install PyInstaller

```powershell
pip install pyinstaller
```

### Step 2: Build the EXE

```powershell
pyinstaller --onefile --console --name "D365ComparisonTool" main.py
```

This creates a single `.exe` file in the `dist` folder.

### Options Explained:
- `--onefile`: Creates a single executable (instead of multiple files)
- `--console`: Shows console window (required for our interactive prompts)
- `--name`: Custom name for the executable

### Step 3: Run the EXE

Navigate to `dist` folder and double-click `D365ComparisonTool.exe`

## Method 2: EXE with Icon (Advanced)

If you have an icon file (`.ico`):

```powershell
pyinstaller --onefile --console --icon=app.ico --name "D365ComparisonTool" main.py
```

## Method 3: Include All Modules Explicitly

If you encounter import errors, use:

```powershell
pyinstaller --onefile --console ^
    --hidden-import=requests ^
    --hidden-import=openpyxl ^
    --hidden-import=auth_manager ^
    --hidden-import=schema_comparison ^
    --hidden-import=excel_generator ^
    --name "D365ComparisonTool" ^
    main.py
```

## Full Build Script (build_exe.bat)

Create this file for easy building:

```batch
@echo off
echo Building D365 Comparison Tool executable...
echo.

REM Install PyInstaller if not present
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM Clean previous builds
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist *.spec del /q *.spec

REM Build the executable
echo.
echo Building executable...
pyinstaller --onefile --console --name "D365ComparisonTool" main.py

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo   Build Complete!
    echo ======================================================================
    echo.
    echo Executable location: dist\D365ComparisonTool.exe
    echo.
    echo You can distribute this single .exe file!
    echo.
) else (
    echo.
    echo Build failed! Check errors above.
    echo.
)

pause
```

Save this as `build_exe.bat` and run it to create your executable.

## Distribution

After building:

1. The `.exe` file will be in the `dist` folder
2. You can distribute ONLY the `.exe` file
3. No Python installation needed on target machines
4. File size will be ~15-30 MB (includes Python runtime)

## Important Notes

### Pros:
- ✓ No Python required on target machine
- ✓ Single file distribution
- ✓ Easy for end users

### Cons:
- ✗ Large file size (~15-30 MB)
- ✗ May trigger antivirus warnings (false positives)
- ✗ Slower startup (unpacks to temp folder)
- ✗ Windows Defender SmartScreen may block unsigned exe

### Antivirus Issues

Some antivirus software flags PyInstaller executables. To resolve:

1. **Code signing**: Sign the .exe with a certificate (costs money)
2. **Whitelist**: Add exception in antivirus
3. **VirusTotal**: Upload to VirusTotal to build reputation
4. **Alternative**: Use the `.bat` launcher instead (simpler, no AV issues)

## Recommended Approach

For internal use, the **`.bat` launcher** (already created) is recommended:
- ✓ No antivirus issues
- ✓ Smaller size
- ✓ Easier to update
- ✓ More transparent

For external distribution to non-technical users, build the `.exe`.

## Testing the EXE

After building:

```powershell
cd dist
.\D365ComparisonTool.exe
```

The application should run exactly like `python main.py`.

## Troubleshooting

**"Failed to execute script"**
- Missing modules: Use `--hidden-import` flags
- Check build output for errors

**Antivirus blocks execution**
- Add exception or get code signing certificate
- Use `.bat` launcher instead

**Import errors at runtime**
- Some modules need explicit inclusion
- Use `--collect-all packagename` flag

**Large file size**
- Normal for PyInstaller (includes Python runtime)
- Use UPX compression: `pip install pyinstaller[upx]`

---

For most users, simply using **`run.bat`** (already created) is the easiest solution!
