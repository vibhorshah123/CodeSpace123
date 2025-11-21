# Dynamics 365 Environment Comparison Tool (Python)

A Python-based interactive command-line tool for comparing Dynamics 365 environments. This tool allows you to compare table schemas, data records with relationships, and generate comprehensive Excel reports.

## Features

- **Interactive Browser-Based Login**: Supports OAuth 2.0 with PKCE for secure authentication (works with MFA and managed devices)
- **Schema Comparison**: Compare table structures between two D365 environments
  - Fields only in source environment
  - Fields only in target environment
  - Fields with differences in properties (type, length, required level, etc.)
  - Matching fields
- **Data Comparison**: Compare actual data records between environments
  - Main record comparison with field-level differences
  - Related records comparison (One-To-Many relationships/subgrids)
  - Automatic relationship discovery
  - Identifies records only in source, only in target, and mismatches
- **Excel Report Generation**: Creates comprehensive Excel reports with multiple sheets:
  - Summary statistics
  - Detailed differences with color coding
  - Related entity comparisons with parent-child hierarchy

## Prerequisites

- Python 3.8 or higher
- Access to two Dynamics 365 environments
- Valid authentication credentials (user account or service principal)

## Installation

### Option 1: Easy Setup (Recommended for Windows)

Simply **double-click** `setup.bat` or run it from command prompt:

```batch
setup.bat
```

This will automatically install all dependencies.

### Option 2: Manual Setup

1. **Navigate to the project directory:**
   ```powershell
   cd c:\Users\v-vibhorshah\source\repos\CodeSpace123\D365PythonComparison
   ```

2. **Install required dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

## Usage

### Option 1: Batch File Launcher (Easiest)

Simply **double-click** `run.bat` or run it from command prompt:

```batch
run.bat
```

### Option 2: Python Command

Start the tool by running the main script:

```powershell
python main.py
```

### Option 3: Standalone EXE (No Python Required)

Build a standalone executable that works without Python:

1. Run `build_exe.bat` to create the executable
2. Find the generated file in `dist\D365ComparisonTool.exe`
3. Copy the `.exe` to any Windows computer (no Python needed!)

See `BUILD_EXE_GUIDE.md` for detailed instructions.

### Step-by-Step Guide

1. **Authentication:**
   - Choose authentication method:
     - Option 1: Username/Password
     - Option 2: Client Credentials (Service Principal)
   - Enter required credentials when prompted

2. **Environment Configuration:**
   - Enter Source Environment URL (e.g., `https://orgname.crm.dynamics.com`)
   - Enter Target Environment URL (e.g., `https://orgname2.crm.dynamics.com`)

3. **Select Comparison Type:**
   - Option 1: Schema Comparison (Compare table structures)
   - Option 2: Data Comparison (Compare records with relationships)
   - Option 3: Flow Comparison (Coming soon)

4. **Schema Comparison:**
   - Enter table logical name (e.g., `contact`, `account`, `mash_customtable`)
   - Review the comparison summary in the console
   - Choose to generate an Excel report

5. **Data Comparison:**
   - Enter table logical name (e.g., `incident`, `account`, `mash_customtable`)
   - Choose comparison scope:
     - Option 1: Main records only
     - Option 2: Main records + related records (One-To-Many relationships/subgrids)
   - Tool will automatically discover and compare related entities
   - Review the comparison summary showing:
     - Total records in each environment
     - Matching records
     - Records only in source/target
     - Field-level mismatches
     - Related entity statistics
   - Generate Excel report with detailed analysis

### Example Session

```
======================================================================
          Dynamics 365 Environment Comparison Tool
======================================================================

Please provide authentication details:
----------------------------------------------------------------------
Auth method (1=Username/Password, 2=Client Credentials): 1
Username (email): user@contoso.com
Password: ********

Enter environment details:
----------------------------------------------------------------------
Source Environment URL: https://contoso-dev.crm.dynamics.com
Target Environment URL: https://contoso-prod.crm.dynamics.com

Authenticating...
✓ Authentication successful!

Select comparison type:
----------------------------------------------------------------------
1. Schema Comparison (Compare table structures)
2. Data Comparison (Coming soon)
3. Flow Comparison (Coming soon)
0. Exit
----------------------------------------------------------------------
Enter your choice: 1

======================================================================
                    Schema Comparison
======================================================================

Enter table logical name: contact

Fetching schema for table 'contact' from both environments...
  Fetching source metadata from https://contoso-dev.crm.dynamics.com...
  Fetching target metadata from https://contoso-prod.crm.dynamics.com...
  Analyzing 245 source fields and 243 target fields...

----------------------------------------------------------------------
COMPARISON SUMMARY
----------------------------------------------------------------------
Table: contact
Source Environment: https://contoso-dev.crm.dynamics.com
Target Environment: https://contoso-prod.crm.dynamics.com

Fields only in Source: 2
Fields only in Target: 0
Fields with differences: 3
Matching fields: 243
----------------------------------------------------------------------

Fields ONLY in Source:
  - mash_newfield
  - mash_testfield

Fields with DIFFERENCES:
  - mash_customfield:
      MaxLength: Source='100' | Target='200'

----------------------------------------------------------------------
Generate Excel report? (y/n): y
Enter output filename (default: schema_comparison.xlsx): contact_comparison.xlsx

✓ Excel report generated: contact_comparison.xlsx
```

## Excel Report Structure

The generated Excel report contains multiple sheets:

1. **Summary**: Overview of comparison results with statistics
2. **Only in Source**: List of fields present only in source environment
3. **Only in Target**: List of fields present only in target environment
4. **Field Differences**: Detailed comparison showing property differences
5. **Matching Fields**: List of fields that match perfectly between environments

## Authentication Methods

### Username/Password (Option 1)
- Uses Resource Owner Password Credentials (ROPC) flow
- Suitable for personal accounts
- Requires MFA to be disabled or configured appropriately

### Client Credentials (Option 2)
- Uses Application authentication
- Requires:
  - Client ID (Application ID)
  - Client Secret
  - Tenant ID
- Recommended for automated scenarios
- Requires app registration in Azure AD with appropriate API permissions

## Project Structure

```
D365PythonComparison/
├── main.py                  # Main entry point and interactive menu
├── auth_manager.py          # Authentication management
├── schema_comparison.py     # Schema comparison logic
├── excel_generator.py       # Excel report generation
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Module Descriptions

### `main.py`
- Entry point for the application
- Handles user interaction and menu navigation
- Orchestrates authentication, comparison, and report generation

### `auth_manager.py`
- Manages OAuth token acquisition
- Supports multiple authentication methods
- Implements token caching

### `schema_comparison.py`
- Fetches table metadata from D365 Web API
- Normalizes field attributes for comparison
- Identifies differences and similarities

### `data_comparison.py`
- Fetches actual data records from D365 environments
- Compares records and identifies mismatches
- Discovers and compares related records (One-To-Many relationships)
- Supports automatic relationship discovery
- Groups child records by parent for subgrid comparison

### `excel_generator.py`
- Creates formatted Excel workbooks
- Applies styling and color coding
- Auto-adjusts column widths
- Generates separate reports for schema and data comparisons

## Troubleshooting

### Authentication Errors

**Problem:** "Authentication failed" or "401 Unauthorized"
- Verify credentials are correct
- Ensure the user/app has access to both environments
- For client credentials, verify API permissions in Azure AD

### API Errors

**Problem:** "Failed to fetch metadata"
- Check environment URLs are correct and accessible
- Verify network connectivity
- Ensure the table logical name is correct

### Permission Issues

**Problem:** "Access denied" errors
- Ensure user/app has appropriate security roles
- System Administrator or System Customizer role typically required for metadata access

## Future Enhancements

Planned features:
- Data comparison (comparing actual records)
- Flow comparison (comparing Power Automate flows)
- Security role comparison
- Solution component comparison
- Batch comparison of multiple tables
- Configuration file support
- Progress bars for long operations

## Related Projects

This tool is inspired by the .NET-based D365 Comparison Function App located at:
```
c:\Users\v-vibhorshah\source\repos\CodeSpace123\D365ComparissionTool\
```

## License

Internal use only.

## Support

For issues or questions, please contact the development team.

---

**Version:** 1.0.0  
**Last Updated:** November 2025
