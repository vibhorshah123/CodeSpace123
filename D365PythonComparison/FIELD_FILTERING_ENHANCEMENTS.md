# Data Comparison Enhancements - Field Filtering & GUID Detection

## Overview

Enhanced the data comparison feature with intelligent field filtering, GUID/lookup mismatch detection, and primary name field matching to provide more meaningful comparisons focused on business data rather than system metadata.

## New Features

### 1. System Field Exclusion List

Automatically excludes system-managed fields from comparison to focus on business-critical data:

#### Excluded Field Categories:

**Ownership & Organization:**
- `owningbusinessunit`
- `owningteam`
- `owninguser`
- `ownerid`

**Audit & Versioning:**
- `versionnumber` ⭐
- `modifiedon` ⭐
- `createdon`
- `modifiedby`
- `createdby`
- `modifiedonbehalfby`
- `createdonbehalfby`
- `overriddencreatedon`

**System Management:**
- `importsequencenumber`
- `timezoneruleversionnumber`
- `utcconversiontimezonecode`
- `processid`
- `stageid`
- `traversedpath`

**Internal Tracking Fields:**
- All `_*_value` fields for excluded lookups (e.g., `_owningbusinessunit_value`)

**Why This Matters:**
- These fields change automatically during system operations
- They differ between environments by design
- Including them creates noise in comparison reports
- Focus on actual business data differences

### 2. GUID/Lookup Field Detection

Automatically identifies and highlights lookup/relationship fields:

#### Detection Methods:

1. **Field Name Pattern:**
   - Fields starting with `_` and ending with `_value` (e.g., `_customerid_value`)
   
2. **Value Pattern:**
   - UUID/GUID format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

#### Benefits:

- **Relationship Mismatches:** Identifies when records point to different related records
- **Migration Issues:** Detects lookup fields that weren't properly migrated
- **Data Integrity:** Ensures related data consistency across environments

#### Example Use Cases:

```
Field: _customerid_value
Source: 12345678-1234-1234-1234-123456789012
Target: 87654321-4321-4321-4321-210987654321

⚠️ This indicates the record points to different customers in each environment!
```

### 3. Primary Name Field Matching

Uses the table's primary name field to find records with same name but different IDs:

#### How It Works:

1. Queries D365 metadata to get `PrimaryNameAttribute` (e.g., `name`, `fullname`, `mash_name`)
2. Builds lookup dictionary by name in addition to ID
3. Identifies records where:
   - Same name exists in both environments
   - But have different primary key IDs

#### Why This Matters:

**Duplicate Detection:**
- Same customer name with different IDs = potential duplicate
- Same account name = possible data quality issue

**Migration Validation:**
- Records that should have migrated but got new IDs
- Helps identify incomplete migrations

**Data Reconciliation:**
- Find records that represent same business entity
- Even if system IDs differ

#### Example Output:

```
⚠ Records with same name but different IDs: 5
(Possible duplicates or migration issues)

Record Name: "Contoso Ltd"
Source ID: abc-123-def-456
Target ID: xyz-789-uvw-012
Status: Same name, different ID - possible duplicate or migration issue
```

## Excel Report Enhancements

### New Sheets Added:

#### 1. Enhanced Field Mismatches Sheet

**Columns:**
- Record ID
- Record Name (primary name field)
- Field Name
- Source Value
- Target Value
- **Type** (GUID/Lookup or Regular)

**Color Coding:**
- 🔴 Red: GUID/Lookup mismatches (relationship issues)
- 🟡 Yellow: Regular field mismatches

#### 2. GUID Mismatches Sheet

Dedicated sheet for lookup/relationship field differences:

**Columns:**
- Record ID
- Record Name
- Field Name
- Source GUID
- Target GUID

**Purpose:**
- Focus on relationship integrity issues
- Easier to spot migration problems
- Helps with data linking validation

#### 3. Name-ID Conflicts Sheet

Records with same primary name but different IDs:

**Columns:**
- Source ID
- Target ID
- Primary Name Field Value
- Status

**Use Cases:**
- Duplicate detection
- Migration validation
- Data quality assessment

### Updated Summary Sheet

**New Metrics:**
- GUID/Lookup Mismatches (count)
- Same Name, Different IDs (count)

**Helps Answer:**
- How many relationship mismatches exist?
- Are there duplicate records across environments?
- What's the data quality status?

## Console Output Enhancements

### New Summary Information:

```
COMPARISON SUMMARY
----------------------------------------------------------------------
Table: account
Source Environment: https://mashtest.crm.dynamics.com
Target Environment: https://mashvnext.crm.dynamics.com

Source records: 150
Target records: 145
Matching records: 130
Records only in Source: 20
Records only in Target: 15
Field mismatches: 45

⚠ GUID/Lookup Mismatches: 12
⚠ Records with same name but different IDs: 5
   (Possible duplicates or migration issues)

Comparing 35 fields (excluding 18 system fields)
```

## Technical Implementation

### 1. System Field Exclusion

```python
def _should_exclude_field(self, field_name: str) -> bool:
    # Check exclusion list (case-insensitive)
    if field_name.lower() in [f.lower() for f in self.SYSTEM_FIELDS_EXCLUSION_LIST]:
        return True
    
    # Exclude OData annotation fields
    if field_name.startswith("@"):
        return True
    
    return False
```

### 2. GUID Detection

```python
def _is_guid_field(self, field_name: str, val1: Any, val2: Any) -> bool:
    # Check field name patterns for lookups
    if field_name.startswith("_") and field_name.endswith("_value"):
        return True
    
    # Check if values look like GUIDs (UUID format)
    guid_pattern = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE)
    
    if val1 and isinstance(val1, str) and guid_pattern.match(val1):
        return True
    
    return False
```

### 3. Primary Name Field Discovery

```python
def _get_primary_name_field(self, environment_url: str, table_logical_name: str, token: str) -> str:
    api_url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')?$select=PrimaryNameAttribute"
    
    response = requests.get(api_url, headers=headers, timeout=30)
    data = response.json()
    return data.get("PrimaryNameAttribute", "name")
```

### 4. Name-Based Matching

```python
# Build lookup dictionaries by primary name
source_by_name = {}
for rec in source_records:
    name = rec.get(primary_name_field)
    if name:
        if name not in source_by_name:
            source_by_name[name] = []
        source_by_name[name].append(rec)

# Check if record with same name exists in target
if source_name and source_name in target_by_name:
    # Record with same name exists but different ID
    name_matches_with_different_ids.append({
        "source_id": record_id,
        "target_id": target_rec.get(pk_field),
        "name": source_name,
        "status": "Same name, different ID - possible duplicate or migration issue"
    })
```

## Customization

### Add Custom Fields to Exclusion List

Edit `data_comparison.py` and add fields to `SYSTEM_FIELDS_EXCLUSION_LIST`:

```python
SYSTEM_FIELDS_EXCLUSION_LIST = [
    # ... existing fields ...
    
    # Add your custom exclusions
    "mycustom_internalfield",
    "mycustom_systemfield",
]
```

### Modify GUID Detection Logic

If you need custom logic for detecting lookup fields:

```python
def _is_guid_field(self, field_name: str, val1: Any, val2: Any) -> bool:
    # Your custom logic here
    if field_name in ["my_special_lookup", "my_reference_field"]:
        return True
    
    # ... existing logic ...
```

## Use Cases

### 1. Post-Migration Validation

**Scenario:** Migrated data from old to new environment

**Use This Feature To:**
- Verify relationships migrated correctly (GUID matching)
- Find records that got new IDs (name matching)
- Identify missing or extra records
- Exclude system timestamps that will differ

**Result:**
- Focus on business data accuracy
- Identify relationship breaks
- Find duplicate records

### 2. Environment Synchronization

**Scenario:** Keep dev/test/prod environments in sync

**Use This Feature To:**
- Compare business data only (exclude system fields)
- Identify configuration differences
- Spot lookup field mismatches
- Track data drift

**Result:**
- Accurate comparison of business logic
- Identify data inconsistencies
- Maintain environment parity

### 3. Duplicate Detection

**Scenario:** Multiple environments with overlapping data

**Use This Feature To:**
- Find records with same name but different IDs
- Identify potential duplicates
- Validate data quality

**Result:**
- Clean up duplicate records
- Improve data integrity
- Reduce data redundancy

### 4. Data Quality Assessment

**Scenario:** Audit data across environments

**Use This Feature To:**
- Focus on custom fields (exclude system fields)
- Identify relationship integrity issues
- Find incomplete records

**Result:**
- Improved data quality
- Better relationship management
- Cleaner datasets

## Best Practices

### 1. Review GUID Mismatches First

Lookup/relationship mismatches can cause:
- Broken workflows
- Incorrect reporting
- User experience issues

**Action:** Fix GUID mismatches before addressing regular field differences

### 2. Investigate Name-ID Conflicts

Records with same name but different IDs may indicate:
- Data migration issues
- Duplicate records
- Integration problems

**Action:** Decide whether to merge, delete, or keep separate

### 3. Add Environment-Specific Exclusions

Some custom fields may be environment-specific:
- Development flags
- Test data markers
- Environment identifiers

**Action:** Add to exclusion list to reduce noise

### 4. Document Intentional Differences

Some GUID mismatches may be expected:
- Different service accounts per environment
- Environment-specific integrations
- Test data vs production data

**Action:** Document these in your comparison reports

## Limitations

1. **Static Exclusion List**: Exclusion list must be manually maintained
2. **Case Sensitivity**: Field names are case-sensitive in D365
3. **Custom Lookup Detection**: May not catch all custom lookup patterns
4. **Name Uniqueness**: Assumes primary name field has reasonable uniqueness

## Future Enhancements

- [ ] Dynamic exclusion list from configuration file
- [ ] Custom field type detection rules
- [ ] Fuzzy name matching (handle spelling variations)
- [ ] Bulk record merging tools
- [ ] Relationship graph visualization
- [ ] Historical comparison tracking

## Summary

These enhancements provide:
✅ **Cleaner Comparisons** - Focus on business data, not system metadata
✅ **Better Insights** - Identify relationship and data quality issues
✅ **Actionable Results** - Clear indication of what needs fixing
✅ **Migration Validation** - Verify data moved correctly
✅ **Duplicate Detection** - Find potential duplicate records

The comparison tool now provides enterprise-grade data validation capabilities! 🎉
