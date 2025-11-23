# Solution Comparison Guide

## Overview
The Solution Comparison feature allows you to compare Dynamics 365/Power Platform solution components between two environments. This helps validate solution migrations, ensure component parity, and identify missing or extra components.

## What is Compared
- **Solution Metadata**: Version, managed/unmanaged status, friendly name
- **Solution Components**: All components within the solution including:
  - Entities (Tables)
  - Attributes (Columns)
  - Forms
  - Views
  - Workflows (Classic Workflows)
  - Power Automate Flows
  - Business Rules
  - Web Resources
  - Plugin Assemblies
  - Plugin Types
  - SDK Message Processing Steps
  - App Modules
  - Canvas Apps
  - Model-driven Apps
  - Environment Variables
  - Connection References
  - And 150+ other component types

## Excel Report Structure
The generated Excel report contains the following sheets:

### 1. Summary
- Solution metadata (name, version, managed status)
- Environment URLs
- Overall statistics:
  - Total components in source
  - Total components in target
  - Common components
  - Components only in source
  - Components only in target

### 2. Component Type Summary
- Breakdown by component type showing:
  - Source count
  - Target count
  - Common count
  - Only in source count
  - Only in target count
- Color-coded:
  - Green: Common components
  - Orange/Yellow: Differences

### 3. Only in Source
- Components present in source but missing in target
- Shows component type, object ID, root behavior
- Highlighted in orange/yellow to indicate missing components

### 4. Only in Target
- Components present in target but not in source
- Shows component type, object ID, root behavior
- Highlighted in blue to indicate extra components

### 5. Common Components
- Components present in both environments
- Shows component type, object ID, root behavior
- Highlighted in light green to indicate parity

## Usage

### Interactive Menu
1. Start the application: `python main.py`
2. Select option **4. Solution Comparison**
3. Enter the solution unique name (e.g., `mysolution`)
4. Wait for comparison to complete
5. Review the console summary
6. Choose whether to generate Excel report

### Example Session
```
Select comparison type:
----------------------------------------------------------------------
1. Schema Comparison (Compare table structures)
2. Data Comparison (Compare records + relationships)
3. Flow Comparison (Compare Power Automate flows)
4. Solution Comparison (Compare solution components)
0. Exit
----------------------------------------------------------------------
Enter your choice: 4

Enter solution unique name (e.g., mysolution): CrmSolution

Comparing solution 'CrmSolution' between environments...

----------------------------------------------------------------------
COMPARISON SUMMARY
----------------------------------------------------------------------
Solution: CrmSolution
Source Environment: https://mashtest.crm.dynamics.com
Target Environment: https://mashvnext.crm.dynamics.com

Source Version: 1.0.0.5
Source Is Managed: No
Target Version: 1.0.0.4
Target Is Managed: No

Source Components: 125
Target Components: 118
Common Components: 115
Only in Source: 10
Only in Target: 3
----------------------------------------------------------------------

Component Type Summary:
----------------------------------------------------------------------
Component Type                 Source     Target     Common
----------------------------------------------------------------------
Entity                         15         15         15
Attribute                      45         43         43
Form                          8          8          8
View                          6          6          6
Workflow                      3          2          2
Web Resource                  18         16         16
Plugin Type                   12         12         12
SDK Message Processing Step   10         8          8
Environment Variable Def      5          5          5
Canvas App                    3          3          3

Components Only in Source (10):
  ⚠ Attribute: guid-1234-5678
  ⚠ Attribute: guid-abcd-efgh
  ⚠ Workflow: guid-flow-0001
  ⚠ Web Resource: guid-webres-01
  ... and 6 more

Generate Excel report? (y/n): y

  Excel report saved to: solution_comparison_CrmSolution_20240315_143022.xlsx

✓ Excel report generated: solution_comparison_CrmSolution_20240315_143022.xlsx
```

## Component Comparison Logic

### Comparison Key
Components are compared using a combination of:
- **Object ID** (GUID): Unique identifier for each component
- **Component Type**: Numeric type code (1=Entity, 29=Workflow, 61=Web Resource, etc.)

### Comparison Process
1. Fetch solution from both environments by unique name
2. Fetch all components for each solution via `solutioncomponents` entity
3. Create tuples of (objectid, componenttype) for each component
4. Compare sets to identify:
   - Components only in source
   - Components only in target
   - Common components (in both)

### Component Types Mapped
The tool maps 160+ component type codes to readable names:
- 1 = Entity
- 2 = Attribute
- 24 = Form
- 26 = View
- 29 = Workflow
- 31 = Relationship
- 60 = Display String
- 61 = Web Resource
- 62 = Web Resource Dependency
- 80 = Plugin Type
- 81 = Plugin Assembly
- 82 = SDK Message Processing Step
- 90 = Plugin Type Statistic
- 91 = Service Endpoint
- 92 = Routing Rule
- 93 = Routing Rule Item
- 95 = App Module Metadata
- 162 = Environment Variable Definition
- 163 = Environment Variable Value
- 300 = Canvas App
- 301 = Connector
- 371 = Connection Reference
- And many more...

## Use Cases

### 1. Solution Migration Validation
After migrating a solution from Dev to Test or Test to Prod:
- Verify all components were deployed
- Identify missing components
- Ensure version numbers match

### 2. Environment Parity Check
Compare solutions across parallel environments:
- Ensure UAT matches Prod
- Validate backup/restore operations
- Check for drift between environments

### 3. Solution Dependency Analysis
Before deploying dependent solutions:
- Verify base solution components exist
- Identify missing prerequisites
- Plan deployment sequence

### 4. Solution Cleanup
Identify components to remove:
- Find extra components in target
- Clean up test components
- Prepare for solution uninstall

### 5. Version Control
Track component changes between versions:
- Compare v1.0 vs v1.1
- Document added/removed components
- Generate change logs

## Benefits
- **Comprehensive**: Covers all 160+ component types
- **Fast**: Direct API queries, no manual inspection
- **Detailed**: Component-level granularity
- **Visual**: Color-coded Excel reports
- **Actionable**: Clearly identifies differences
- **Automated**: Consistent, repeatable process

## Limitations
- **Component Content**: Only compares component presence, not internal details
- **Dependencies**: Does not analyze component dependencies
- **Customizations**: Does not detect changes within component definitions
- **Versioning**: Only shows version numbers, not detailed change history
- **Permissions**: Requires read access to both environments and solution entities

## Tips for Best Results
1. **Use Unique Names**: Solution unique name is case-sensitive
2. **Check Permissions**: Ensure access to solutions and solutioncomponents entities
3. **Consider Managed vs Unmanaged**: Managed solutions may have different component sets
4. **Version Tracking**: Compare same versions for accurate results
5. **Regular Checks**: Run comparisons before and after deployments
6. **Export Reports**: Keep Excel reports for audit trail
7. **Review Missing Components**: Investigate why components are missing - intentional or error?

## Troubleshooting

### Solution Not Found
- Verify solution unique name (case-sensitive)
- Check solution exists in both environments
- Ensure you have read permissions to solutions entity

### No Components Returned
- Solution may be empty (default solution)
- Check permissions to solutioncomponents entity
- Verify solution ID is correct

### Large Component Count
- Large solutions may take longer to fetch
- Consider comparing subsets if needed
- Use pagination for very large solutions

### Component Type "Unknown"
- Some component types may not be mapped
- Check component type code in data
- Report unmapped types for future enhancement

## API Details

### Endpoints Used
```
GET /api/data/v9.2/solutions?$filter=uniquename eq '{solution_name}'
GET /api/data/v9.2/solutioncomponents?$filter=_solutionid_value eq '{solution_id}'
```

### Component Attributes
- `objectid`: GUID of the component
- `componenttype`: Numeric type code
- `rootcomponentbehavior`: Component inclusion behavior
- `solutionid`: Parent solution reference

## Technical Details
- **Module**: `src/solution_comparison.py`
- **Excel Generator**: `src/excel_generator.py`
- **API Version**: D365 Web API v9.2
- **Authentication**: OAuth 2.0 with PKCE
- **Pagination**: Automatic for large component sets
- **Normalization**: Component types mapped to friendly names

## Future Enhancements
- [ ] Component content comparison (deep diff)
- [ ] Dependency tree visualization
- [ ] Multi-solution comparison
- [ ] Historical version tracking
- [ ] Component change detection
- [ ] Export to JSON/CSV formats
- [ ] Solution deployment automation
