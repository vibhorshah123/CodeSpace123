# Quick Start Guide

## 1. Setup (One-time)

**Simply double-click** `setup.bat` in the D365PythonComparison folder!

Or run from command prompt:

```batch
setup.bat
```

This will install all required dependencies automatically.

## 2. Run the Tool

**Simply double-click** `run.bat`!

Or run from command prompt:

```batch
run.bat
```

Or if you prefer Python directly:

```powershell
python main.py
```

## 3. Choose Authentication

**Option 1: Username/Password** (Quick & Easy)
- Best for testing
- Uses your D365 credentials
- MFA may need to be disabled

**Option 2: Client Credentials** (Recommended for Automation)
- Best for production use
- Requires Azure AD App Registration
- More secure for repeated use

### Setting up Client Credentials (Option 2)

1. Go to Azure Portal > Azure Active Directory > App Registrations
2. Click "New registration"
3. Give it a name (e.g., "D365 Comparison Tool")
4. Register the application
5. Note the **Client ID** and **Tenant ID**
6. Go to "Certificates & secrets" > New client secret
7. Note the **Client Secret** (shown only once!)
8. Go to "API permissions" > Add permission
9. Select "Dynamics CRM" > Delegated permissions > "user_impersonation"
10. Grant admin consent
11. In D365, create an Application User with System Administrator role

## 4. Run Schema Comparison

1. Enter your source environment URL (e.g., `https://contoso-dev.crm.dynamics.com`)
2. Enter your target environment URL (e.g., `https://contoso-prod.crm.dynamics.com`)
3. Select option `1` for Schema Comparison
4. Enter table name (e.g., `contact`, `account`, `mash_customtable`)
5. Review results
6. Generate Excel report when prompted

## 5. View Results

The Excel report will contain:
- **Summary**: Overview with statistics
- **Only in Source**: Fields missing in target
- **Only in Target**: New fields in target
- **Field Differences**: Property changes
- **Matching Fields**: Identical fields

## Common Table Names

- Standard tables: `contact`, `account`, `lead`, `opportunity`
- Custom tables: Usually start with prefix like `mash_`, `new_`, `cr###_`
- Activity tables: `email`, `phonecall`, `task`, `appointment`

## Troubleshooting

**"Authentication failed"**
- Check credentials are correct
- Verify environment URLs are accessible
- Try pinging the URL in browser

**"Failed to fetch metadata"**
- Check table name spelling (must be logical name, not display name)
- Verify you have permissions in both environments
- Ensure network connectivity

**"No module named 'requests'"**
- Re-run `.\setup.ps1` to install dependencies

## Tips

1. **Find table logical name**: In D365, go to Settings > Customizations > Customize the System > Entities. The "Name" column shows the logical name.

2. **Batch comparisons**: You can run the tool multiple times. Each comparison is independent.

3. **Save your reports**: Use descriptive filenames like `contact_dev_vs_prod_2025-11-20.xlsx`

4. **Before deployments**: Always run comparison before moving solutions between environments

## Next Steps

After schema comparison works, future features will include:
- Data comparison (compare records)
- Flow comparison
- Batch mode for multiple tables
- Automated scheduled comparisons

---

Need help? Check README.md for detailed documentation.
