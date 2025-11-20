# Data Comparison Guide

## Overview

The Data Comparison feature allows you to compare actual data records between two Dynamics 365 environments, including related records (subgrids) through One-To-Many relationships.

## Features

### Main Record Comparison

- **Field-Level Comparison**: Compares all fields in the main table
- **Record Identification**: Identifies records that:
  - Exist only in source environment
  - Exist only in target environment
  - Have field value mismatches
  - Match perfectly between environments

### Relationship Support

The tool supports comparing related records across three relationship types:

1. **One-To-Many Relationships** (Implemented)
   - Parent-to-child relationships (subgrids)
   - Automatic relationship discovery
   - Groups child records by parent record
   - Identifies differences in child record counts

2. **Many-To-One Relationships** (Coming Soon)
   - Child-to-parent lookup relationships

3. **Many-To-Many Relationships** (Coming Soon)
   - Association records between entities

## Usage

### Step 1: Select Data Comparison

From the main menu, select option **2. Data Comparison**

### Step 2: Enter Table Name

Enter the logical name of the table you want to compare:

```
Enter table logical name (e.g., contact, account, mash_customtable): incident
```

### Step 3: Choose Comparison Scope

Select what to compare:

```
Comparison Options:
  1. Compare main records only
  2. Compare main records + related records (One-To-Many relationships/subgrids)

Select option [1-2]: 2
```

### Step 4: Wait for Processing

The tool will:
1. Fetch records from source environment
2. Fetch records from target environment
3. Discover relationships (if option 2 selected)
4. Fetch related records for all relationships
5. Compare all data and identify differences

### Step 5: Review Summary

View the comparison summary in console:

```
COMPARISON SUMMARY
----------------------------------------------------------------------
Table: incident
Source Environment: https://mashtest.crm.dynamics.com
Target Environment: https://mashvnext.crm.dynamics.com

Source records: 150
Target records: 145
Matching records: 130
Records only in Source: 20
Records only in Target: 15
Field mismatches: 45

Related Entities Compared: 3
  - annotation: 234 source, 245 target
      Differences found in 12 parent records
  - incidentresolution: 89 source, 87 target
      Differences found in 5 parent records
  - activitypointer: 456 source, 459 target
      Differences found in 8 parent records
```

### Step 6: Generate Excel Report

Choose to generate a detailed Excel report:

```
Generate Excel report? (y/n): y
Enter output filename (default: data_comparison.xlsx): incident_comparison.xlsx

✓ Excel report generated: incident_comparison.xlsx
```

## Excel Report Structure

The generated Excel workbook contains multiple sheets:

### 1. Summary Sheet

- Report metadata (table name, environments, timestamp)
- Comparison statistics
- Related entities summary

### 2. Only in Source Sheet

- Complete records that exist only in source environment
- All fields included

### 3. Only in Target Sheet

- Complete records that exist only in target environment
- All fields included

### 4. Field Mismatches Sheet

- Record ID
- Field name
- Source value
- Target value
- Color-coded for easy identification

### 5. Matching Records Sheet

- List of record IDs that match perfectly between environments

### 6. Related Entity Sheets (One sheet per entity)

For each related entity compared:
- Entity name and lookup field
- Total record counts (source and target)
- Differences by parent record:
  - Parent ID
  - Count of records only in source
  - Count of records only in target
  - Status indicator (color-coded)

## Best Practices

### 1. Performance Considerations

- **Large tables**: The tool fetches ALL active records by default
- **Consider filtering**: For very large tables, consider adding custom filters in the code
- **Relationship limits**: By default, only the first 5 One-To-Many relationships are compared

### 2. Data Volume

- Excel has a limit of ~1 million rows per sheet
- For very large datasets, consider:
  - Comparing specific record subsets
  - Using filters to reduce data volume
  - Breaking comparison into multiple runs

### 3. Security Permissions

Ensure your account has:
- **Read permission** on the table
- **Read permission** on related entities
- Access to both environments

### 4. Active Records Only

By default, the tool only compares **active records** (statecode = 0). To include inactive records, modify the `only_active` parameter in `data_comparison.py`.

## Advanced Usage

### Filtering Specific Records

To compare only specific records, you can modify the filter in `data_comparison.py`:

```python
# Example: Compare only high priority incidents
result = comparison.compare_table_data(
    source_env,
    target_env,
    "incident",
    fields=None,
    filter_query="prioritycode eq 1",  # High priority only
    include_relationships=True
)
```

### Selecting Specific Fields

To compare only specific fields:

```python
# Example: Compare only title and description fields
result = comparison.compare_table_data(
    source_env,
    target_env,
    "incident",
    fields=["incidentid", "title", "description"],
    include_relationships=False
)
```

### Custom Relationship Configuration

To compare specific relationships instead of auto-discovery:

```python
specific_rels = [
    {
        "ChildEntityLogicalName": "annotation",
        "ChildParentFieldLogicalName": "objectid"
    },
    {
        "ChildEntityLogicalName": "incidentresolution",
        "ChildParentFieldLogicalName": "incidentid"
    }
]

result = comparison.compare_table_data(
    source_env,
    target_env,
    "incident",
    fields=None,
    include_relationships=True,
    specific_relationships=specific_rels
)
```

## Comparison Logic

### Record Matching

Records are matched using the **primary key field** (e.g., `incidentid` for `incident` table).

### Field Value Comparison

- **Null values**: Treated as equal to empty strings
- **String comparison**: Case-insensitive
- **Type differences**: Converted to string before comparison

### Child Record Comparison

For One-To-Many relationships:
1. Fetch all child records for all parent IDs
2. Group child records by parent ID
3. Compare child record counts per parent
4. Identify missing/extra child records

## Troubleshooting

### Issue: "Failed to fetch records"

**Cause**: Permission issues or invalid table name

**Solution**:
- Verify table logical name is correct
- Check security roles include Read permission
- Ensure environment URL is accessible

### Issue: No relationships discovered

**Cause**: Table has no One-To-Many relationships or they're system-only

**Solution**:
- Verify relationships exist in Dataverse
- Check if relationships are valid for data access
- Some system relationships may be filtered out

### Issue: Excel report too large

**Cause**: Too many records or relationships

**Solution**:
- Compare smaller subsets of data
- Use field selection to reduce columns
- Limit relationship comparison
- Consider comparing in batches

### Issue: Slow performance

**Cause**: Large data volume or many relationships

**Solution**:
- Reduce comparison scope (main records only)
- Limit number of relationships
- Add filters to reduce record count
- Run during off-peak hours

## Limitations

1. **One-To-Many relationships only**: Many-To-One and Many-To-Many coming soon
2. **Active records only**: By default, inactive records are excluded
3. **Relationship limit**: Only first 5 relationships compared by default
4. **No incremental comparison**: Always compares full dataset
5. **No change tracking**: Doesn't track when changes occurred

## Coming Soon

- [ ] Many-To-One relationship support
- [ ] Many-To-Many relationship support
- [ ] Custom filter UI in main menu
- [ ] Progress bars for long operations
- [ ] Incremental comparison (only changed records)
- [ ] Change tracking with timestamps
- [ ] Batch processing for very large datasets
- [ ] Custom comparison rules per field type

## Examples

### Example 1: Compare Account Records Only

```
Select comparison type: 2
Enter table logical name: account
Select option [1-2]: 1

Result: Compares only account records, no related entities
```

### Example 2: Compare Incidents with Related Records

```
Select comparison type: 2
Enter table logical name: incident
Select option [1-2]: 2

Result: Compares incident records + all related entities (notes, activities, etc.)
```

### Example 3: Compare Custom Entity

```
Select comparison type: 2
Enter table logical name: mash_customtable
Select option [1-2]: 2

Result: Compares custom entity records + any related entities
```
