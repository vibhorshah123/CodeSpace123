# D365 Python Comparison Tool - Project Structure

## Overview
This document explains the organization of the D365 Python Comparison Tool project.

## Folder Structure

```
D365PythonComparison/
├── src/                          # Source code modules
│   ├── __init__.py               # Python package initialization
│   ├── auth_manager.py           # OAuth 2.0 authentication with PKCE
│   ├── schema_comparison.py      # Table metadata comparison
│   ├── data_comparison.py        # Data record comparison with relationships
│   └── excel_generator.py        # Excel report generation
│
├── docs/                         # Documentation
│   ├── README.md                 # Main project documentation
│   ├── QUICKSTART.md             # Getting started guide
│   ├── BUILD_EXE_GUIDE.md        # PyInstaller executable guide
│   ├── DATA_COMPARISON_GUIDE.md  # Data comparison feature guide
│   ├── FIELD_FILTERING_ENHANCEMENTS.md  # System field exclusion guide
│   ├── START_HERE.txt            # Simple first-time user guide
│   └── PROJECT_STRUCTURE.md      # This file
│
├── main.py                       # Application entry point
├── requirements.txt              # Python dependencies
├── config.sample.json            # Sample configuration file
├── .gitignore                    # Git ignore patterns
├── run.bat                       # Windows launcher
├── setup.bat                     # Dependency installer
└── build_exe.bat                 # PyInstaller build script
```

## Key Components

### Source Code (src/)
- **auth_manager.py**: Handles OAuth 2.0 authentication using Authorization Code Flow with PKCE. Opens browser for user login and captures tokens via local HTTP server on port 8765.

- **schema_comparison.py**: Compares table metadata (attributes, display names, types, required levels) between two D365 environments. Generates Excel reports with color-coded differences.

- **data_comparison.py**: Compares actual data records between environments using GUID-based matching. Features:
  - System field exclusion (33 fields including modifiedon, createdby, ownerid)
  - One-To-Many relationship comparison (subgrids)
  - GUID mismatch detection for lookup fields
  - Primary name field matching

- **excel_generator.py**: Creates formatted Excel workbooks with multiple sheets showing comparison results. Uses color coding and clear GUID-based terminology.

### Documentation (docs/)
- **README.md**: Comprehensive project documentation
- **QUICKSTART.md**: Step-by-step setup instructions
- **BUILD_EXE_GUIDE.md**: Instructions for creating standalone executable
- **DATA_COMPARISON_GUIDE.md**: Detailed data comparison feature documentation
- **FIELD_FILTERING_ENHANCEMENTS.md**: System field exclusion and GUID detection documentation
- **START_HERE.txt**: Simple text guide for beginners

### Root Files
- **main.py**: Interactive menu-driven application entry point
- **requirements.txt**: Python package dependencies (requests, openpyxl, tqdm)
- **config.sample.json**: Template for user configuration
- **run.bat**: Convenience launcher for Windows
- **setup.bat**: Automated dependency installation
- **build_exe.bat**: PyInstaller executable builder

## Import Structure

All modules in the `src/` folder are imported using the `src.` prefix:

```python
from src.auth_manager import AuthManager
from src.schema_comparison import SchemaComparison
from src.data_comparison import DataComparison
from src.excel_generator import ExcelGenerator
```

## Running the Application

### From Python
```bash
python main.py
```

### Using Batch File (Windows)
```bash
run.bat
```

### As Standalone Executable
```bash
build_exe.bat          # Build once
dist\D365ComparisonTool.exe  # Run anywhere
```

## Development Guidelines

1. **Adding New Features**: Create new modules in `src/` folder
2. **Documentation**: Add guides to `docs/` folder
3. **Testing**: Run `python main.py` after changes to verify imports
4. **Building**: Use `build_exe.bat` to create distributable executable

## Benefits of This Structure

- **Separation of Concerns**: Source code separate from documentation
- **Python Best Practices**: Proper package structure with `__init__.py`
- **Maintainability**: Clear organization makes code easier to understand
- **Scalability**: Easy to add new modules and documentation
- **Professional**: Standard Python project layout
