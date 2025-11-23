# Flow Comparison Feature

## Overview
Power Automate flow comparison functionality has been added to the D365 Python Comparison Tool, based on the C# implementation from D365ComparissionTool.

## Features

### Flow Comparison
- **Fetch Flows**: Retrieves cloud flows from Dataverse workflows entity (category 5)
- **Normalization**: 
  - Removes environment-specific data (connectionReferences, timestamps, etc.)
  - Masks GUIDs with `<GUID>` placeholder
  - Masks environment URLs with `<ENV_HOST>` placeholder
  - Sorts JSON keys alphabetically for consistent comparison
- **Hash Comparison**: Computes SHA256 hash of normalized flow definitions
- **Detailed Diff**: Tracks added, removed, and changed properties
- **Action-Level Analysis**: Identifies differences in individual flow actions

### Comparison Modes
1. **All Flows**: Compare all flows between two environments
2. **Single Flow**: Compare a specific flow by name

### Excel Report Generation

The tool generates comprehensive Excel reports with multiple sheets:

#### 1. Summary Sheet
- Environment details (using actual env names like 'mash', 'mashvnext')
- Total flow counts
- Identical flows count
- Different flows count
- Missing in target count
- Flows with errors count

#### 2. Identical Flows Sheet
- List of flows with identical definitions
- Green highlighting for success

#### 3. Different Flows Sheet
- Flow name
- Hash comparison (first 16 chars)
- Count of added, removed, and changed paths
- Color-coded for visibility

#### 4. Action Differences Sheet
- Flow name
- Action name
- Status (ADDED, REMOVED, CHANGED)
- Property path
- Source and target values
- Detailed property-level changes

#### 5. Missing in Target Sheet
- Flows present in source but not in target
- Flow ID and name
- Warning highlighting

#### 6. Errors Sheet
- Flows that encountered processing errors
- Source and target error messages

### File Naming
Auto-generated filenames follow the pattern:
- Single flow: `flow_comparison_{flowname}_{timestamp}.xlsx`
- All flows: `flow_comparison_all_{timestamp}.xlsx`
- Timestamp format: `YYYYMMDD_HHMMSS`

## Usage

### From Main Menu
```
Select comparison type:
----------------------------------------------------------------------
1. Schema Comparison (Compare table structures)
2. Data Comparison (Compare records + relationships)
3. Flow Comparison (Compare Power Automate flows)
0. Exit
----------------------------------------------------------------------
Enter your choice: 3
```

### Flow Comparison Process
1. Select comparison mode (all flows or specific flow)
2. If specific flow, enter the flow name
3. Tool fetches flows from both environments
4. Normalizes and compares flow definitions
5. Displays summary in console
6. Optionally generates Excel report

### Example Output
```
COMPARISON SUMMARY
----------------------------------------------------------------------
Source Environment: https://mash.crm.dynamics.com
Target Environment: https://mashvnext.crm.dynamics.com

Source flows: 25
Target flows: 23
Identical flows: 18
Different flows: 5
Missing in Target: 2
Flows with errors: 0
----------------------------------------------------------------------

Identical Flows (18):
  ✓ Approval Flow
  ✓ Notification Handler
  ✓ Data Sync Process
  ✓ Email Automation
  ✓ Status Update Flow
  ... and 13 more

Different Flows (5):
  ⚠ Customer Onboarding
  ⚠ Invoice Processing
  ⚠ Lead Assignment
  ⚠ Order Fulfillment
  ⚠ Record Update Flow

Missing in Target (2):
  ✗ New Feature Flow
  ✗ Test Automation Flow
```

## Technical Details

### Normalization Logic
The flow comparison removes the following from definitions:
- `connectionReferences` - Environment-specific connection info
- `runtimeConfiguration` - Runtime settings
- `lastModified`, `createdTime`, `modifiedTime` - Timestamps
- `etag` - Version tags
- `trackedProperties` - Tracking metadata
- Any property ending with `id` or `Id`
- Any property starting with `connection`

### API Endpoint
```
https://{environment}/api/data/v9.2/workflows?$select=workflowid,name,clientdata&$filter=category eq 5
```

Category 5 specifically targets cloud flows (Power Automate).

### Authentication
Uses the same OAuth authentication manager as schema and data comparison, supporting:
- Interactive browser login (OAuth with PKCE)
- Service principal (client credentials)

## Benefits

1. **Cross-Environment Validation**: Ensure flows are deployed correctly
2. **Change Detection**: Identify unintended modifications
3. **Documentation**: Generate reports for audit purposes
4. **Migration Support**: Validate flow migrations between environments
5. **Version Control**: Track flow definition changes over time

## Limitations

- Compares flow definitions only (not run history or analytics)
- Requires read access to workflows entity in both environments
- Large flow definitions may take longer to process
- Connection references are masked and not compared

## Future Enhancements

Possible future improvements:
- Flow run history comparison
- Connection reference validation
- Solution-aware flow comparison
- Batch comparison with filtering
- Export normalized definitions to JSON
