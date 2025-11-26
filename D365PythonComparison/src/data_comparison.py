"""
Data Comparison Module
Compares actual data records between two Dynamics 365 environments.
Simplified: record-by-record comparison using the table's **Primary Name** field only (no GUID matching).
Relationships are NOT compared.
"""

import requests
from typing import Dict, List, Any, Optional
from src.auth_manager import AuthManager


class DataComparison:
    """Handles data comparison operations between D365 environments"""
    
    # System fields to exclude from comparison
    SYSTEM_FIELDS_EXCLUSION_LIST = [
        "owningbusinessunit","owningteam","owninguser","ownerid",
        "versionnumber","modifiedon","createdon","modifiedby","createdby",
        "modifiedonbehalfby","createdonbehalfby","overriddencreatedon",
        "importsequencenumber","timezoneruleversionnumber","utcconversiontimezonecode",
        "processid","stageid","traversedpath",
        "_owningbusinessunit_value","_owningteam_value","_owninguser_value","_ownerid_value",
        "_modifiedby_value","_createdby_value","_modifiedonbehalfby_value","_createdonbehalfby_value"
    ]

    def __init__(self, auth_manager: AuthManager):
        self.auth_manager = auth_manager

    def fetch_records(self, environment_url: str, table_logical_name: str, 
                     fields: Optional[List[str]] = None, 
                     filter_query: Optional[str] = None,
                     only_active: bool = True) -> List[Dict[str, Any]]:
        token = self.auth_manager.get_token(environment_url)
        base_url = environment_url.rstrip('/')
        entity_set_name = self._get_entity_set_name(environment_url, table_logical_name, token)
        api_url = f"{base_url}/api/data/v9.2/{entity_set_name}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=\"*\""
        }
        params = []
        if fields:
            params.append(f"$select={','.join(fields)}")
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
        
        print(f"     Fetched {len(records)} records total")
        return records

    def _should_exclude_field(self, field_name: str) -> bool:
        if field_name.lower() in [f.lower() for f in self.SYSTEM_FIELDS_EXCLUSION_LIST]:
            return True
        if field_name.startswith("@"):
            return True
        if field_name.endswith("@OData.Community.Display.V1.FormattedValue"):
            return True
        return False

    def _get_primary_name_field(self, environment_url: str, table_logical_name: str, token: str) -> str:
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
            return "name"

    def _get_entity_set_name(self, environment_url: str, table_logical_name: str, token: str) -> str:
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
            return table_logical_name + "s"

    def compare_table_data(self, source_url: str, target_url: str, 
                          table_logical_name: str,
                          fields: Optional[List[str]] = None) -> Dict[str, Any]:
        print(f"\n  Comparing data for table: {table_logical_name}")
        print(f"  Source: {source_url}")
        print(f"  Target: {target_url}")
        print(f"  Match Mode: PRIMARY NAME")
        print()
        
        print("  Fetching source records...")
        source_records = self.fetch_records(source_url, table_logical_name, fields)
        print("  Fetching target records...")
        target_records = self.fetch_records(target_url, table_logical_name, fields)
        
        token = self.auth_manager.get_token(source_url)
        primary_name_field = self._get_primary_name_field(source_url, table_logical_name, token)
        print(f"  Primary name field: {primary_name_field}")
        
        def build_name_lookup(records: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
            lookup = {}
            for rec in records:
                name = rec.get(primary_name_field)
                if name:
                    key = str(name).lower().strip()
                    lookup.setdefault(key, []).append(rec)
            return lookup
        
        source_by_name = build_name_lookup(source_records)
        target_by_name = build_name_lookup(target_records)
        
        print(f"\n  Comparing by Primary Name: {len(source_by_name)} unique names in source vs {len(target_by_name)} in target...")
        result = self._compare_by_name(source_by_name, target_by_name, primary_name_field, source_records, target_records)
        
        result.update({
            "table_name": table_logical_name,
            "source_url": source_url,
            "target_url": target_url,
            "source_record_count": len(source_records),
            "target_record_count": len(target_records),
            "primary_name_field": primary_name_field,
            "match_mode": "name"
        })
        
        self._print_comparison_summary(result)
        return result

    def _compare_by_name(self, source_by_name: Dict, target_by_name: Dict,
                         primary_name_field: str,
                         source_records: List, target_records: List) -> Dict[str, Any]:
        only_in_source = []
        only_in_target = []
        mismatches = []
        matching_records = []
        name_matched_records = []
        duplicate_names_source = []
        duplicate_names_target = []
        
        for name_key, records in source_by_name.items():
            if len(records) > 1:
                duplicate_names_source.append({"name": records[0].get(primary_name_field), "count": len(records)})
        for name_key, records in target_by_name.items():
            if len(records) > 1:
                duplicate_names_target.append({"name": records[0].get(primary_name_field), "count": len(records)})
        
        pk_guess = None
        if source_records:
            pk_guess = next((f for f in source_records[0].keys() if f.endswith("id")), None)
        
        all_fields = set()
        for rec in source_records + target_records:
            all_fields.update(rec.keys())
        compared_fields = sorted([f for f in all_fields if not self._should_exclude_field(f)])
        compared_fields_no_pk = [f for f in compared_fields if f != pk_guess]
        print(f"  Comparing {len(compared_fields_no_pk)} fields (excluding system fields and primary key)")
        
        processed_names = set()
        
        for name_key, source_recs in source_by_name.items():
            processed_names.add(name_key)
            if name_key not in target_by_name:
                only_in_source.extend(source_recs)
                continue
            
            target_recs = target_by_name[name_key]
            source_rec = source_recs[0]
            target_rec = target_recs[0]
            
            source_id = source_rec.get(pk_guess)
            target_id = target_rec.get(pk_guess)
            record_name = source_rec.get(primary_name_field)
            
            record_mismatches = []
            for field in compared_fields_no_pk:
                source_val = source_rec.get(field)
                target_val = target_rec.get(field)
                if not self._values_equal(source_val, target_val):
                    is_guid = self._is_guid_field(field, source_val, target_val)
                    record_mismatches.append({
                        "record_id": source_id,
                        "target_record_id": target_id,
                        "record_name": record_name,
                        "field_name": field,
                        "source_value": source_val,
                        "target_value": target_val,
                        "is_guid": is_guid,
                        "matched_by": "name"
                    })
            
            if record_mismatches:
                mismatches.extend(record_mismatches)
            else:
                matching_records.append(source_id)
            
            if source_id != target_id:
                name_matched_records.append({
                    "source_id": source_id,
                    "target_id": target_id,
                    "name": record_name,
                    "has_differences": len(record_mismatches) > 0
                })
        
        for name_key, target_recs in target_by_name.items():
            if name_key not in processed_names:
                only_in_target.extend(target_recs)
        
        return {
            "only_in_source": only_in_source,
            "only_in_target": only_in_target,
            "mismatches": mismatches,
            "matching_records": matching_records,
            "compared_fields": compared_fields_no_pk,
            "name_matched_records": name_matched_records,
            "duplicate_names_source": duplicate_names_source,
            "duplicate_names_target": duplicate_names_target,
        }

    def _is_guid_field(self, field_name: str, val1: Any, val2: Any) -> bool:
        if field_name.startswith("_") and field_name.endswith("_value"):
            return True
        import re
        guid_pattern = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE)
        if val1 and isinstance(val1, str) and guid_pattern.match(val1):
            return True
        if val2 and isinstance(val2, str) and guid_pattern.match(val2):
            return True
        return False

    def _values_equal(self, val1: Any, val2: Any) -> bool:
        if val1 is None or val1 == "":
            return val2 is None or val2 == ""
        if val2 is None or val2 == "":
            return val1 is None or val1 == ""
        return str(val1).lower() == str(val2).lower()

    def _print_comparison_summary(self, result: Dict):
        print(f"\n   Comparison complete! (Mode: PRIMARY NAME)")
        print(f"    - Matching names (identical): {len(result['matching_records'])}")
        print(f"    - Names with attribute mismatches: {len(set([m['record_id'] for m in result['mismatches'] if m.get('record_id')]))}")
        print(f"      (Total field mismatches: {len(result['mismatches'])})")
        print(f"    - Names only in source: {len(result['only_in_source'])}")
        print(f"    - Names only in target: {len(result['only_in_target'])}")
        if result.get('name_matched_records'):
            print(f"    - ℹ Records matched by name (different GUIDs observed): {len(result['name_matched_records'])}")
        if result.get('duplicate_names_source'):
            print(f"    -  Duplicate names in source: {len(result['duplicate_names_source'])}")
        if result.get('duplicate_names_target'):
            print(f"    -  Duplicate names in target: {len(result['duplicate_names_target'])}")
