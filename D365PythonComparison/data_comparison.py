"""
Data Comparison Module
Compares actual data records between two Dynamics 365 environments
Includes support for related records (subgrids) with One-To-Many, Many-To-One, and Many-To-Many relationships
"""

import requests
from typing import Dict, List, Any, Optional, Tuple
from auth_manager import AuthManager


class DataComparison:
    """Handles data comparison operations between D365 environments"""
    
    # System fields to exclude from comparison
    SYSTEM_FIELDS_EXCLUSION_LIST = [
        # Ownership and organization
        "owningbusinessunit",
        "owningteam",
        "owninguser",
        "ownerid",
        
        # Audit and versioning
        "versionnumber",
        "modifiedon",
        "createdon",
        "modifiedby",
        "createdby",
        "modifiedonbehalfby",
        "createdonbehalfby",
        "overriddencreatedon",
        
        # System management
        "importsequencenumber",
        "timezoneruleversionnumber",
        "utcconversiontimezonecode",
        "processid",
        "stageid",
        "traversedpath",
        
        # Internal tracking
        "_owningbusinessunit_value",
        "_owningteam_value",
        "_owninguser_value",
        "_ownerid_value",
        "_modifiedby_value",
        "_createdby_value",
        "_modifiedonbehalfby_value",
        "_createdonbehalfby_value"
    ]
    
    def __init__(self, auth_manager: AuthManager):
        """
        Initialize data comparison
        
        Args:
            auth_manager: AuthManager instance for authentication
        """
        self.auth_manager = auth_manager
    
    def fetch_records(self, environment_url: str, table_logical_name: str, 
                     fields: Optional[List[str]] = None, 
                     filter_query: Optional[str] = None,
                     only_active: bool = True) -> List[Dict[str, Any]]:
        """
        Fetch records from D365 environment
        
        Args:
            environment_url: URL of the D365 environment
            table_logical_name: Logical name of the table
            fields: List of fields to retrieve (None = all fields)
            filter_query: OData filter query
            only_active: Whether to filter only active records (statecode=0)
            
        Returns:
            List of record dictionaries
        """
        token = self.auth_manager.get_token(environment_url)
        base_url = environment_url.rstrip('/')
        
        # Get entity set name from metadata
        entity_set_name = self._get_entity_set_name(environment_url, table_logical_name, token)
        
        api_url = f"{base_url}/api/data/v9.2/{entity_set_name}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=\"*\""
        }
        
        # Build query parameters
        params = []
        
        if fields:
            params.append(f"$select={','.join(fields)}")
        
        # Build filter
        filters = []
        if only_active:
            filters.append("statecode eq 0")
        if filter_query:
            filters.append(filter_query)
        
        if filters:
            params.append(f"$filter={' and '.join(filters)}")
        
        query_string = "&".join(params) if params else ""
        full_url = f"{api_url}?{query_string}" if query_string else api_url
        
        records = []
        next_link = full_url
        
        print(f"    Fetching records from {table_logical_name}...")
        
        while next_link:
            try:
                response = requests.get(next_link, headers=headers, timeout=60)
                response.raise_for_status()
                
                data = response.json()
                records.extend(data.get("value", []))
                
                next_link = data.get("@odata.nextLink")
                
                if len(records) % 100 == 0:
                    print(f"      Fetched {len(records)} records...")
                    
            except requests.RequestException as e:
                error_detail = ""
                if hasattr(e, 'response') and e.response is not None:
                    try:
                        error_detail = f"\n{e.response.text[:500]}"
                    except:
                        pass
                raise Exception(f"Failed to fetch records: {str(e)}{error_detail}")
        
        print(f"    ✓ Fetched {len(records)} records total")
        return records
    
    def _should_exclude_field(self, field_name: str) -> bool:
        """
        Check if field should be excluded from comparison
        
        Args:
            field_name: Name of the field to check
            
        Returns:
            True if field should be excluded, False otherwise
        """
        # Check exclusion list (case-insensitive)
        if field_name.lower() in [f.lower() for f in self.SYSTEM_FIELDS_EXCLUSION_LIST]:
            return True
        
        # Exclude OData annotation fields
        if field_name.startswith("@"):
            return True
        
        # Exclude formatted value fields
        if field_name.endswith("@OData.Community.Display.V1.FormattedValue"):
            return True
        
        return False
    
    def _get_primary_name_field(self, environment_url: str, table_logical_name: str, token: str) -> str:
        """
        Get the primary name field for a table
        
        Args:
            environment_url: URL of the D365 environment
            table_logical_name: Logical name of the table
            token: Auth token
            
        Returns:
            Primary name field logical name
        """
        base_url = environment_url.rstrip('/')
        api_url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')?$select=PrimaryNameAttribute"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }
        
        try:
            response = requests.get(api_url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            return data.get("PrimaryNameAttribute", "name")
        except:
            # Fallback to common primary name field
            return "name"
    
    def _get_entity_set_name(self, environment_url: str, table_logical_name: str, token: str) -> str:
        """Get entity set name from metadata"""
        base_url = environment_url.rstrip('/')
        api_url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')?$select=EntitySetName"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }
        
        try:
            response = requests.get(api_url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            return data.get("EntitySetName", table_logical_name + "s")
        except:
            # Fallback to simple pluralization
            return table_logical_name + "s"
    
    def discover_relationships(self, environment_url: str, table_logical_name: str) -> Dict[str, List[Dict[str, Any]]]:
        """
        Discover all relationships for a table
        
        Args:
            environment_url: URL of the D365 environment
            table_logical_name: Logical name of the table
            
        Returns:
            Dictionary with relationship types and their details
        """
        token = self.auth_manager.get_token(environment_url)
        base_url = environment_url.rstrip('/')
        
        relationships = {
            "OneToMany": [],
            "ManyToOne": [],
            "ManyToMany": []
        }
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }
        
        # Get One-To-Many relationships
        print(f"    Discovering One-To-Many relationships...")
        url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')/OneToManyRelationships?$select=ReferencingEntity,ReferencingAttribute,SchemaName"
        try:
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            relationships["OneToMany"] = data.get("value", [])
            print(f"      Found {len(relationships['OneToMany'])} One-To-Many relationships")
        except:
            pass
        
        # Get Many-To-One relationships
        print(f"    Discovering Many-To-One relationships...")
        url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')/ManyToOneRelationships?$select=ReferencedEntity,ReferencedAttribute,ReferencingAttribute,SchemaName"
        try:
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            relationships["ManyToOne"] = data.get("value", [])
            print(f"      Found {len(relationships['ManyToOne'])} Many-To-One relationships")
        except:
            pass
        
        # Get Many-To-Many relationships
        print(f"    Discovering Many-To-Many relationships...")
        url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')/ManyToManyRelationships?$select=Entity1LogicalName,Entity2LogicalName,IntersectEntityName,SchemaName"
        try:
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            relationships["ManyToMany"] = data.get("value", [])
            print(f"      Found {len(relationships['ManyToMany'])} Many-To-Many relationships")
        except:
            pass
        
        return relationships
    
    def compare_table_data(self, source_url: str, target_url: str, 
                          table_logical_name: str,
                          fields: Optional[List[str]] = None,
                          include_relationships: bool = True,
                          specific_relationships: Optional[List[Dict[str, str]]] = None) -> Dict[str, Any]:
        """
        Compare table data between two environments
        
        Args:
            source_url: Source environment URL
            target_url: Target environment URL
            table_logical_name: Logical name of the table to compare
            fields: Specific fields to compare (None = all fields)
            include_relationships: Whether to include related records comparison
            specific_relationships: Specific relationships to compare (None = auto-discover)
            
        Returns:
            Dictionary containing comparison results
        """
        print(f"\n  Comparing data for table: {table_logical_name}")
        print(f"  Source: {source_url}")
        print(f"  Target: {target_url}")
        print()
        
        # Fetch source records
        print("  Fetching source records...")
        source_records = self.fetch_records(source_url, table_logical_name, fields)
        
        # Fetch target records
        print("  Fetching target records...")
        target_records = self.fetch_records(target_url, table_logical_name, fields)
        
        # Get primary key field and primary name field
        pk_field = table_logical_name + "id"
        token = self.auth_manager.get_token(source_url)
        primary_name_field = self._get_primary_name_field(source_url, table_logical_name, token)
        
        print(f"  Primary key field (GUID): {pk_field}")
        print(f"  Primary name field: {primary_name_field}")
        print(f"  ℹ Comparison is based on GUID matching only")
        
        # Build lookup dictionaries by ID (GUID-based)
        source_dict = {rec.get(pk_field): rec for rec in source_records if rec.get(pk_field)}
        target_dict = {rec.get(pk_field): rec for rec in target_records if rec.get(pk_field)}
        
        # Build lookup dictionaries by primary name (for finding records with same name but different ID)
        source_by_name = {}
        for rec in source_records:
            name = rec.get(primary_name_field)
            if name:
                if name not in source_by_name:
                    source_by_name[name] = []
                source_by_name[name].append(rec)
        
        target_by_name = {}
        for rec in target_records:
            name = rec.get(primary_name_field)
            if name:
                if name not in target_by_name:
                    target_by_name[name] = []
                target_by_name[name].append(rec)
        
        print(f"\n  Analyzing records by GUID: {len(source_dict)} source vs {len(target_dict)} target...")
        
        # Find differences (GUID-based comparison)
        only_in_source = []  # GUIDs present in source but not in target
        only_in_target = []  # GUIDs present in target but not in source
        mismatches = []  # Same GUID but different attribute values
        matching_records = []  # Same GUID with identical attributes
        guid_mismatches = []  # Lookup field mismatches (related record GUIDs differ)
        name_matches_with_different_ids = []
        
        # Get all field names from both sets and filter out system fields
        all_fields = set()
        for rec in source_records + target_records:
            all_fields.update(rec.keys())
        
        # Filter fields: exclude system fields and OData annotations
        compared_fields = sorted([f for f in all_fields if not self._should_exclude_field(f)])
        excluded_count = len(all_fields) - len(compared_fields)
        
        print(f"  Comparing {len(compared_fields)} fields (excluding {excluded_count} system fields)")
        print(f"  ℹ Excluded: modifiedon, modifiedby, createdon, createdby, ownerid, versionnumber, etc.")
        
        # Compare common records (by ID)
        for record_id, source_rec in source_dict.items():
            if record_id not in target_dict:
                # Check if a record with same primary name exists in target
                source_name = source_rec.get(primary_name_field)
                if source_name and source_name in target_by_name:
                    # Record with same name exists but different ID
                    for target_rec in target_by_name[source_name]:
                        name_matches_with_different_ids.append({
                            "source_id": record_id,
                            "target_id": target_rec.get(pk_field),
                            "name": source_name,
                            "status": "Same name, different ID - possible duplicate or migration issue"
                        })
                
                only_in_source.append(source_rec)
                continue
            
            target_rec = target_dict[record_id]
            record_mismatches = []
            record_guid_mismatches = []
            
            for field in compared_fields:
                source_val = source_rec.get(field)
                target_val = target_rec.get(field)
                
                if not self._values_equal(source_val, target_val):
                    # Check if this is a GUID field (lookup)
                    is_guid_mismatch = self._is_guid_field(field, source_val, target_val)
                    
                    mismatch_detail = {
                        "record_id": record_id,
                        "record_name": source_rec.get(primary_name_field, ""),
                        "field_name": field,
                        "source_value": source_val,
                        "target_value": target_val,
                        "is_guid": is_guid_mismatch
                    }
                    
                    record_mismatches.append(mismatch_detail)
                    
                    if is_guid_mismatch:
                        record_guid_mismatches.append(mismatch_detail)
            
            if record_mismatches:
                mismatches.extend(record_mismatches)
                if record_guid_mismatches:
                    guid_mismatches.extend(record_guid_mismatches)
            else:
                matching_records.append(record_id)
        
        # Find records only in target
        for record_id in target_dict:
            if record_id not in source_dict:
                # Check if a record with same primary name exists in source
                target_name = target_dict[record_id].get(primary_name_field)
                if target_name and target_name in source_by_name:
                    # Already handled in the source loop
                    pass
                
                only_in_target.append(target_dict[record_id])
        
        result = {
            "table_name": table_logical_name,
            "source_url": source_url,
            "target_url": target_url,
            "source_record_count": len(source_records),
            "target_record_count": len(target_records),
            "only_in_source": only_in_source,
            "only_in_target": only_in_target,
            "mismatches": mismatches,
            "matching_records": matching_records,
            "compared_fields": compared_fields,
            "guid_mismatches": guid_mismatches,
            "name_matches_with_different_ids": name_matches_with_different_ids,
            "primary_name_field": primary_name_field,
            "child_comparisons": {}
        }
        
        # Compare related records if requested
        if include_relationships:
            print(f"\n  Discovering relationships for {table_logical_name}...")
            relationships = self.discover_relationships(source_url, table_logical_name)
            
            # Process One-To-Many relationships (subgrids)
            if relationships["OneToMany"]:
                print(f"\n  Processing {len(relationships['OneToMany'])} One-To-Many relationships...")
                result["child_comparisons"] = self._compare_one_to_many_relationships(
                    source_url, target_url, table_logical_name,
                    relationships["OneToMany"], source_dict, target_dict
                )
        
        print(f"\n  ✓ GUID-based comparison complete!")
        print(f"    - Matching GUIDs (identical): {len(matching_records)}")
        print(f"    - GUIDs with attribute mismatches: {len(set([m['record_id'] for m in mismatches]))}")
        print(f"      (Total field mismatches: {len(mismatches)}, including {len(guid_mismatches)} lookup field mismatches)")
        print(f"    - GUIDs only in source: {len(only_in_source)}")
        print(f"    - GUIDs only in target: {len(only_in_target)}")
        if name_matches_with_different_ids:
            print(f"    - ⚠ Records with same name but different GUIDs: {len(name_matches_with_different_ids)}")
        
        return result
    
    def _compare_one_to_many_relationships(self, source_url: str, target_url: str,
                                          parent_table: str, relationships: List[Dict[str, Any]],
                                          source_parent_dict: Dict, target_parent_dict: Dict) -> Dict[str, Any]:
        """Compare One-To-Many related records (subgrids)"""
        child_results = {}
        
        for rel in relationships[:5]:  # Limit to first 5 relationships to avoid too much data
            child_entity = rel.get("ReferencingEntity")
            lookup_field = rel.get("ReferencingAttribute")
            
            if not child_entity or not lookup_field:
                continue
            
            # Skip system relationships
            if child_entity.startswith("msdyn_") or child_entity in ["annotation", "activitypointer"]:
                continue
            
                print(f"\n    Comparing child entity: {child_entity} (via {lookup_field})")
                print(f"      Using parent GUIDs for matching...")
            
            try:
                # Fetch child records from source (using parent GUIDs)
                parent_ids_source = list(source_parent_dict.keys())
                parent_ids_target = list(target_parent_dict.keys())
                
                if not parent_ids_source and not parent_ids_target:
                    continue
                
                # Fetch child records for all parent IDs
                source_children = self._fetch_related_records(
                    source_url, child_entity, lookup_field, parent_ids_source
                )
                
                target_children = self._fetch_related_records(
                    target_url, child_entity, lookup_field, parent_ids_target
                )
                
                # Group by parent
                source_by_parent = self._group_by_parent(source_children, lookup_field)
                target_by_parent = self._group_by_parent(target_children, lookup_field)
                
                # Compare
                child_pk = child_entity + "id"
                child_mismatches = []
                
                for parent_id in set(source_by_parent.keys()) | set(target_by_parent.keys()):
                    source_child_records = source_by_parent.get(parent_id, [])
                    target_child_records = target_by_parent.get(parent_id, [])
                    
                    source_child_dict = {r.get(child_pk): r for r in source_child_records if r.get(child_pk)}
                    target_child_dict = {r.get(child_pk): r for r in target_child_records if r.get(child_pk)}
                    
                    # Find differences in child records
                    only_in_source_count = len(set(source_child_dict.keys()) - set(target_child_dict.keys()))
                    only_in_target_count = len(set(target_child_dict.keys()) - set(source_child_dict.keys()))
                    
                    if only_in_source_count > 0 or only_in_target_count > 0:
                        child_mismatches.append({
                            "parent_id": parent_id,
                            "only_in_source_count": only_in_source_count,
                            "only_in_target_count": only_in_target_count
                        })
                
                child_results[child_entity] = {
                    "lookup_field": lookup_field,
                    "source_total": len(source_children),
                    "target_total": len(target_children),
                    "differences": child_mismatches
                }
                
                print(f"      ✓ Source: {len(source_children)} records, Target: {len(target_children)} records")
                
            except Exception as e:
                print(f"      ✗ Error comparing {child_entity}: {str(e)}")
                continue
        
        return child_results
    
    def _fetch_related_records(self, environment_url: str, child_entity: str, 
                               lookup_field: str, parent_ids: List[str]) -> List[Dict[str, Any]]:
        """Fetch related child records for given parent IDs"""
        if not parent_ids:
            return []
        
        # Build filter for parent IDs (in batches to avoid URL length limits)
        batch_size = 20
        all_records = []
        
        for i in range(0, len(parent_ids), batch_size):
            batch = parent_ids[i:i+batch_size]
            filter_parts = [f"{lookup_field} eq {pid}" for pid in batch]
            filter_query = " or ".join(filter_parts)
            
            try:
                records = self.fetch_records(
                    environment_url, 
                    child_entity, 
                    fields=None,  # Get all fields
                    filter_query=filter_query,
                    only_active=True
                )
                all_records.extend(records)
            except:
                continue
        
        return all_records
    
    def _group_by_parent(self, records: List[Dict[str, Any]], lookup_field: str) -> Dict[str, List[Dict[str, Any]]]:
        """Group child records by parent ID"""
        grouped = {}
        for record in records:
            parent_id = record.get(lookup_field)
            if parent_id:
                if parent_id not in grouped:
                    grouped[parent_id] = []
                grouped[parent_id].append(record)
        return grouped
    
    def _is_guid_field(self, field_name: str, val1: Any, val2: Any) -> bool:
        """Check if field appears to be a GUID/lookup field"""
        # Check field name patterns for lookups
        if field_name.startswith("_") and field_name.endswith("_value"):
            return True
        
        # Check if values look like GUIDs (UUID format)
        import re
        guid_pattern = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE)
        
        if val1 and isinstance(val1, str) and guid_pattern.match(val1):
            return True
        if val2 and isinstance(val2, str) and guid_pattern.match(val2):
            return True
        
        return False
    
    def _values_equal(self, val1: Any, val2: Any) -> bool:
        """Compare two values for equality, handling nulls and type differences"""
        # Treat None and empty string as equal
        if val1 is None or val1 == "":
            return val2 is None or val2 == ""
        if val2 is None or val2 == "":
            return val1 is None or val1 == ""
        
        # String comparison (case-insensitive)
        return str(val1).lower() == str(val2).lower()
