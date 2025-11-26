# D365 Python Comparison Tool

## Description

The **D365 Python Comparison Tool** is a powerful command-line utility designed to compare configurations, data, and components across different Dynamics 365 environments. Built with Python, it provides comprehensive comparison capabilities with automated Excel report generation, making it easy to validate deployments, track changes, and ensure environment parity.

The tool leverages Dynamics 365 Web API to perform deep comparisons across multiple dimensions - from table schemas and data records to Power Automate flows and complete solution components. It supports secure OAuth 2.0 authentication (including MFA) and generates professional, color-coded Excel reports for easy analysis and documentation.

---

## Use Cases

### 1. **Pre-Deployment Validation**
Before moving solutions from Dev to Test or Test to Production, validate that all customizations, fields, and configurations exist in the target environment to prevent deployment failures.

### 2. **Post-Deployment Verification**
After solution imports or data migrations, verify that all components were successfully deployed and that data integrity is maintained across environments.

### 3. **Environment Drift Detection**
Regularly audit UAT, Pre-Prod, and Production environments to identify unauthorized changes, missing customizations, or configuration drift.

### 4. **Data Migration Quality Assurance**
Compare data records between source and target environments during migration projects to ensure completeness and accuracy, including related child records.

### 5. **Solution Component Tracking**
Track which components (entities, forms, workflows, web resources, plugins, canvas apps, etc.) are included in solutions and identify missing or extra components across environments.

### 6. **Power Automate Flow Validation**
Compare flow definitions between environments to ensure business logic consistency and detect unintended changes in automation.

### 7. **Compliance & Audit Documentation**
Generate comprehensive comparison reports for compliance audits, change management documentation, and environment certification processes.

### 8. **Troubleshooting Environment Issues**
Quickly identify configuration differences when troubleshooting issues that work in one environment but not another.

### 9. **Development Team Coordination**
Enable multiple developers to validate their changes against shared environments and identify conflicts before merging customizations.

### 10. **Release Management**
Create detailed change logs by comparing environments before and after releases, documenting exactly what was modified.

---

## What's in it for you

The utility performs comprehensive scans across your Dynamics 365 environments and generates detailed comparison reports in Excel format with the following capabilities:

| **Feature** | **Minimum Permission Required** | **Authentication** | **Output** |
|------------|--------------------------------|-------------------|-----------|
| **Schema Comparison** | System Customizer | OAuth 2.0 (Browser Login or Service Principal) | Excel report with field differences, property changes, and metadata comparison |
| **Data Comparison** | Read access on target tables | OAuth 2.0 (Browser Login or Service Principal) | Excel report with record differences, field mismatches, GUID comparisons, and relationship data |
| **Flow Comparison** | Basic User with flow access | OAuth 2.0 (Browser Login or Service Principal) | Excel report with flow definitions, action-level differences, and hash-based comparison |
| **Solution Comparison** | System Customizer | OAuth 2.0 (Browser Login or Service Principal) | Excel report with 160+ component types, component-level differences, and version tracking |
| **Related Records** | Read access on parent & child tables | OAuth 2.0 (Browser Login or Service Principal) | Excel report with One-to-Many relationship comparisons and subgrid data |
| **Automated Reports** | Same as above | Same as above | Auto-generated timestamped Excel files with color-coded differences |

### Key Benefits

✅ **Time Savings**: Automated comparisons that would take hours manually now complete in minutes  
✅ **Accuracy**: Eliminates manual comparison errors with systematic, comprehensive analysis  
✅ **Visibility**: Clear, color-coded Excel reports make differences immediately obvious  
✅ **Audit Trail**: Timestamped reports provide documentation for compliance and change management  
✅ **Environment Names**: Uses actual environment names (not generic "Source/Target") for clarity  
✅ **System Field Exclusion**: Automatically excludes 35+ system fields to focus on business data  
✅ **GUID-Based Matching**: Reliable record comparison using unique identifiers  
✅ **Relationship Support**: Compares parent and child records through One-to-Many relationships  
✅ **Flow Normalization**: Intelligent comparison that ignores environment-specific GUIDs and timestamps  
✅ **Comprehensive Coverage**: 160+ solution component types supported  
✅ **Flexible Authentication**: Supports both interactive login (with MFA) and service principals  
✅ **No Cloud Costs**: Runs locally on your machine, no Azure resources required  
✅ **Open Source**: Full visibility into comparison logic and customization capability  

### What You Get

📊 **Excel Reports** with multiple sheets:
- Summary statistics
- Items only in source environment (missing in target)
- Items only in target environment (extra in target)
- Field/property/component differences with before/after values
- Matching items (confirmation of parity)
- Relationship comparisons (for data)
- Action-level diffs (for flows)
- Component type breakdown (for solutions)

🎨 **Color-Coded Visualization**:
- 🔴 Red/Pink: Missing items or critical differences
- 🟡 Yellow/Orange: Property differences or mismatches
- 🟢 Green: Matching items (validation of parity)
- 🔵 Blue: Informational items

📈 **Comprehensive Metrics**:
- Total counts (source vs target)
- Difference counts and percentages
- Field-level change details
- Component-level granularity
- Relationship statistics

---

## Steps to Use

1. **Download the repository** and run `setup.bat` to install dependencies automatically.
2. **Execute `run.bat`** and provide environment names, authentication (browser login or service principal), and select the comparison type (Schema/Data/Flow/Solution).
3. **Review the auto-generated Excel report** with color-coded differences, summary statistics, and detailed comparison results saved as `{type}_{name}_{timestamp}.xlsx`.

---

## Quick Start Examples

### Schema Comparison
```
Select comparison type: 1
Enter table logical name: contact
Generate Excel report? y

Output: schema_comparison_contact_20251123_143022.xlsx
```

### Data Comparison with Relationships
```
Select comparison type: 2
Enter table logical name: account
Comparison Options:
  1. Compare main records only
  2. Compare main records + related records
Select option: 2

Output: data_comparison_account_20251123_145530.xlsx
```

### Flow Comparison
```
Select comparison type: 3
Enter flow name: My Approval Flow
Generate Excel report? y

Output: flow_comparison_MyApprovalFlow_20251123_150145.xlsx
```

### Solution Comparison
```
Select comparison type: 4
Enter solution unique name: CrmSolution
Generate Excel report? y

Output: solution_comparison_CrmSolution_20251123_152233.xlsx
```

---

## Support & Feedback

**For any queries, feedback, or issues:**

📧 **Contact:** v-vibhorshah@microsoft.com  
💬 **Teams:** @Vibhor Shah | MAQ Software  
📁 **Repository:** CodeSpace123/D365PythonComparison  
📖 **Documentation:** See `docs/` folder for comprehensive guides

**Available Documentation:**
- `QUICKSTART.md` - Quick start guide
- `README.md` - Complete documentation
- `DATA_COMPARISON_GUIDE.md` - Data comparison details
- `FLOW_COMPARISON_GUIDE.md` - Flow comparison details
- `SOLUTION_COMPARISON_GUIDE.md` - Solution comparison details
- `D365_Comparison_Tool_How_To_Guide.docx` - Complete Word guide

**Contributing:**
We welcome feedback, bug reports, and feature requests. Please reach out via Teams or email.

---

## Technical Details

**Technology Stack:**
- Python 3.8+
- Dynamics 365 Web API v9.2
- OAuth 2.0 with PKCE
- OpenPyXL for Excel generation
- Requests library for HTTP

**Security:**
- Secure OAuth 2.0 authentication
- Supports MFA and managed devices
- Token caching for performance
- No credentials stored on disk

**Performance:**
- Parallel API calls where possible
- Pagination for large datasets
- Progress indicators with tqdm
- Efficient memory management

**Compatibility:**
- Windows 7 or later
- Any D365 environment (Online)
- Power Platform environments
- Dataverse environments

---

**Version:** 1.0.0  
**Last Updated:** November 2025  
**License:** Internal use only

---

*This tool is developed and maintained by the MAQ Software D365 team to streamline environment management and ensure deployment quality across Dynamics 365 projects.*
