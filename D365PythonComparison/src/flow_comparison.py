"""
Flow Comparison Module
Compares Power Automate flows between two Dynamics 365 environments
Based on the C# implementation from D365ComparissionTool
"""

import requests
import hashlib
import json
import re
from typing import Dict, List, Any, Optional, Tuple
from src.auth_manager import AuthManager


class FlowComparison:
    """Handles Power Automate flow comparison operations between D365 environments"""
    
    # Keys to ignore during flow normalization for comparison
    DEFAULT_IGNORE_KEYS = [
        "connectionReferences",
        "runtimeConfiguration",
        "lastModified",
        "createdTime",
        "modifiedTime",
        "etag",
        "trackedProperties",
        "workflowid",
        "flowid"
    ]
    
    def __init__(self, auth_manager: AuthManager):
        """
        Initialize flow comparison
        
        Args:
            auth_manager: Authentication manager instance
        """
        self.auth_manager = auth_manager
        self.ignore_keys = self.DEFAULT_IGNORE_KEYS.copy()
    
    def fetch_flows(self, base_url: str, filter_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Fetch flows from Dataverse workflows entity
        
        Args:
            base_url: Base URL for the environment (e.g., https://org.crm.dynamics.com)
            filter_name: Optional flow name to filter by
            
        Returns:
            List of flow dictionaries with workflowid, name, and clientdata
        """
        token = self.auth_manager.get_token(base_url)
        
        # Normalize URL
        if not base_url.startswith("https://"):
            base_url = f"https://{base_url}"
        base_url = base_url.rstrip('/')
        
        # Build API endpoint - category 5 = cloud flows
        api_url = f"{base_url}/api/data/v9.2/workflows?$select=workflowid,name,clientdata&$filter=category eq 5"
        
        if filter_name:
            # Escape single quotes in name
            escaped_name = filter_name.replace("'", "''")
            api_url = f"{base_url}/api/data/v9.2/workflows?$select=workflowid,name,clientdata&$filter=category eq 5 and name eq '{escaped_name}'"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "User-Agent": "D365PythonComparison/1.0"
        }
        
        flows = []
        page = 0
        
        print(f"  Fetching flows from {base_url}...")
        
        while api_url:
            page += 1
            try:
                response = requests.get(api_url, headers=headers, timeout=60)
                
                if response.status_code == 429 or response.status_code >= 500:
                    # Transient error - retry with backoff
                    import time
                    wait_time = min(5, 0.25 * page)
                    print(f"    Status {response.status_code}, retrying in {wait_time}s...")
                    time.sleep(wait_time)
                    continue
                
                response.raise_for_status()
                data = response.json()
                
                # Handle single result or array
                if "value" in data:
                    flows.extend(data["value"])
                elif "workflowid" in data:
                    flows.append(data)
                
                # Check for next page
                api_url = data.get("@odata.nextLink")
                
                if api_url:
                    print(f"    Page {page}: {len(flows)} flows fetched so far...")
                
            except requests.exceptions.RequestException as e:
                print(f"    Error fetching flows: {str(e)}")
                raise
        
        print(f"  ✓ Total flows fetched: {len(flows)}")
        return flows
    
    def normalize_flow_definition(self, flow_json: Dict[str, Any], env_host: str) -> str:
        """
        Normalize flow definition for comparison by:
        - Removing ignore keys
        - Masking GUIDs
        - Masking environment-specific URLs
        - Sorting keys alphabetically
        
        Args:
            flow_json: Flow definition JSON
            env_host: Environment host for masking
            
        Returns:
            Normalized canonical JSON string
        """
        def should_ignore(key: str) -> bool:
            """Check if key should be ignored"""
            if key in self.ignore_keys:
                return True
            # Ignore keys ending with 'id' or 'Id'
            if re.match(r".*[iI]d$", key):
                return True
            # Ignore keys starting with 'connection'
            if key.lower().startswith("connection"):
                return True
            return False
        
        def mask_guids(text: str) -> str:
            """Replace GUIDs with placeholder"""
            return re.sub(
                r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}',
                '<GUID>',
                text
            )
        
        def mask_host_urls(text: str) -> str:
            """Replace environment host with placeholder"""
            if env_host:
                return text.replace(env_host, '<ENV_HOST>')
            return text
        
        def traverse(obj):
            """Recursively normalize object"""
            if isinstance(obj, dict):
                # Sort keys and filter ignored ones
                normalized = {}
                for key in sorted(obj.keys()):
                    if should_ignore(key):
                        continue
                    value = traverse(obj[key])
                    normalized[key] = value
                return normalized
            elif isinstance(obj, list):
                return [traverse(item) for item in obj]
            elif isinstance(obj, str):
                # Mask GUIDs and URLs in strings
                text = mask_guids(obj)
                text = mask_host_urls(text)
                return text
            else:
                return obj
        
        normalized = traverse(flow_json)
        
        # Serialize to canonical JSON (sorted, no whitespace)
        canonical = json.dumps(normalized, sort_keys=True, separators=(',', ':'))
        return canonical
    
    def compute_hash(self, canonical_json: str) -> str:
        """
        Compute SHA256 hash of canonical JSON
        
        Args:
            canonical_json: Normalized canonical JSON string
            
        Returns:
            Hex string of SHA256 hash
        """
        return hashlib.sha256(canonical_json.encode('utf-8')).hexdigest()
    
    def build_snapshot(self, flow: Dict[str, Any], env_host: str) -> Dict[str, Any]:
        """
        Build a flow snapshot with normalized definition and hash
        
        Args:
            flow: Raw flow data from API
            env_host: Environment host
            
        Returns:
            Snapshot dictionary
        """
        flow_id = flow.get("workflowid", "")
        flow_name = flow.get("name", "")
        clientdata = flow.get("clientdata")
        
        snapshot = {
            "flow_id": flow_id,
            "name": flow_name,
            "environment": env_host,
            "canonical_json": "",
            "hash": "",
            "error": None
        }
        
        if not clientdata:
            snapshot["error"] = "clientdata missing or empty"
            return snapshot
        
        try:
            # Parse clientdata if it's a string
            if isinstance(clientdata, str):
                flow_json = json.loads(clientdata)
            else:
                flow_json = clientdata
            
            # Normalize and hash
            canonical = self.normalize_flow_definition(flow_json, env_host)
            flow_hash = self.compute_hash(canonical)
            
            snapshot["canonical_json"] = canonical
            snapshot["hash"] = flow_hash
            
        except Exception as e:
            snapshot["error"] = str(e)
        
        return snapshot
    
    def compute_diff(self, json_a: str, json_b: str) -> Dict[str, Any]:
        """
        Compute detailed differences between two JSON strings
        
        Args:
            json_a: First JSON string
            json_b: Second JSON string
            
        Returns:
            Dictionary with added, removed, and changed paths
        """
        obj_a = json.loads(json_a)
        obj_b = json.loads(json_b)
        
        added = []
        removed = []
        changed = []
        
        def compare_elements(a, b, path="$"):
            """Recursively compare elements"""
            if type(a) != type(b):
                changed.append({
                    "path": path,
                    "source_value": str(a),
                    "target_value": str(b)
                })
                return
            
            if isinstance(a, dict):
                keys_a = set(a.keys())
                keys_b = set(b.keys())
                
                for key in keys_a - keys_b:
                    removed.append(f"{path}.{key}")
                
                for key in keys_b - keys_a:
                    added.append(f"{path}.{key}")
                
                for key in keys_a & keys_b:
                    compare_elements(a[key], b[key], f"{path}.{key}")
            
            elif isinstance(a, list):
                max_len = max(len(a), len(b))
                for i in range(max_len):
                    if i >= len(a):
                        added.append(f"{path}[{i}]")
                    elif i >= len(b):
                        removed.append(f"{path}[{i}]")
                    else:
                        compare_elements(a[i], b[i], f"{path}[{i}]")
            
            else:
                if a != b:
                    changed.append({
                        "path": path,
                        "source_value": str(a),
                        "target_value": str(b)
                    })
        
        compare_elements(obj_a, obj_b)
        
        return {
            "added": added,
            "removed": removed,
            "changed": changed
        }
    
    def extract_actions(self, flow_json: str) -> Optional[Dict[str, Any]]:
        """
        Extract actions from flow definition
        
        Args:
            flow_json: Flow JSON string
            
        Returns:
            Dictionary of action names to action definitions
        """
        try:
            obj = json.loads(flow_json)
            
            # Try properties.definition.actions
            if "properties" in obj and "definition" in obj["properties"]:
                definition = obj["properties"]["definition"]
                if "actions" in definition:
                    return definition["actions"]
            
            # Try definition.actions
            if "definition" in obj and "actions" in obj["definition"]:
                return obj["definition"]["actions"]
            
            return None
        except:
            return None
    
    def compute_action_differences(self, json_a: str, json_b: str) -> List[Dict[str, Any]]:
        """
        Compute action-level differences between flows
        
        Args:
            json_a: First flow JSON
            json_b: Second flow JSON
            
        Returns:
            List of action differences
        """
        actions_a = self.extract_actions(json_a) or {}
        actions_b = self.extract_actions(json_b) or {}
        
        differences = []
        
        # Removed actions
        for name in set(actions_a.keys()) - set(actions_b.keys()):
            differences.append({
                "action_name": name,
                "status": "removed",
                "changed_properties": []
            })
        
        # Added actions
        for name in set(actions_b.keys()) - set(actions_a.keys()):
            differences.append({
                "action_name": name,
                "status": "added",
                "changed_properties": []
            })
        
        # Changed actions
        for name in set(actions_a.keys()) & set(actions_b.keys()):
            action_a_json = json.dumps(actions_a[name], sort_keys=True)
            action_b_json = json.dumps(actions_b[name], sort_keys=True)
            
            if action_a_json != action_b_json:
                # Compute detailed diff for this action
                diff = self.compute_diff(action_a_json, action_b_json)
                changed_props = []
                
                for change in diff["changed"]:
                    changed_props.append({
                        "path": change["path"],
                        "source_value": change["source_value"],
                        "target_value": change["target_value"]
                    })
                
                # Include added/removed as property changes
                for path in diff["added"]:
                    changed_props.append({
                        "path": path,
                        "source_value": "(missing)",
                        "target_value": "(added)"
                    })
                
                for path in diff["removed"]:
                    changed_props.append({
                        "path": path,
                        "source_value": "(removed)",
                        "target_value": "(missing)"
                    })
                
                differences.append({
                    "action_name": name,
                    "status": "changed",
                    "changed_properties": changed_props
                })
        
        return differences
    
    def compare_flows(self, 
                     source_url: str,
                     target_url: str,
                     flow_name: Optional[str] = None,
                     include_diff_details: bool = True) -> Dict[str, Any]:
        """
        Compare flows between two environments
        
        Args:
            source_url: Source environment URL
            target_url: Target environment URL
            flow_name: Optional specific flow name to compare
            include_diff_details: Include detailed diff information
            
        Returns:
            Dictionary containing comparison results
        """
        print(f"\nFetching flows from source environment...")
        source_flows = self.fetch_flows(source_url, flow_name)
        
        print(f"\nFetching flows from target environment...")
        target_flows = self.fetch_flows(target_url, flow_name)
        
        # Extract host names
        source_host = source_url.replace("https://", "").replace("http://", "").split('/')[0]
        target_host = target_url.replace("https://", "").replace("http://", "").split('/')[0]
        
        print(f"\nBuilding flow snapshots...")
        source_snapshots = {}
        for flow in source_flows:
            snapshot = self.build_snapshot(flow, source_host)
            source_snapshots[snapshot["name"]] = snapshot
        
        target_snapshots = {}
        for flow in target_flows:
            snapshot = self.build_snapshot(flow, target_host)
            target_snapshots[snapshot["name"]] = snapshot
        
        print(f"  ✓ Source snapshots: {len(source_snapshots)}")
        print(f"  ✓ Target snapshots: {len(target_snapshots)}")
        
        # Compare flows
        print(f"\nComparing flows...")
        comparisons = []
        identical_flows = []
        non_identical_flows = []
        missing_in_target = []
        errors = []
        
        for name, source_snap in source_snapshots.items():
            target_snap = target_snapshots.get(name)
            
            if not target_snap:
                missing_in_target.append(name)
                comparisons.append({
                    "name": name,
                    "source": source_snap,
                    "target": None,
                    "identical": False,
                    "status": "missing_in_target",
                    "diff": None,
                    "action_differences": None
                })
                continue
            
            # Check for errors
            if source_snap.get("error") or target_snap.get("error"):
                errors.append(name)
                comparisons.append({
                    "name": name,
                    "source": source_snap,
                    "target": target_snap,
                    "identical": False,
                    "status": "error",
                    "diff": None,
                    "action_differences": None
                })
                continue
            
            # Compare hashes
            identical = source_snap["hash"] == target_snap["hash"]
            
            diff = None
            action_diffs = None
            
            if not identical and include_diff_details:
                diff = self.compute_diff(
                    source_snap["canonical_json"],
                    target_snap["canonical_json"]
                )
                action_diffs = self.compute_action_differences(
                    source_snap["canonical_json"],
                    target_snap["canonical_json"]
                )
            
            comparisons.append({
                "name": name,
                "source": source_snap,
                "target": target_snap,
                "identical": identical,
                "status": "identical" if identical else "different",
                "diff": diff,
                "action_differences": action_diffs
            })
            
            if identical:
                identical_flows.append(name)
            else:
                non_identical_flows.append(name)
        
        print(f"  ✓ Comparison complete")
        
        return {
            "source_url": source_url,
            "target_url": target_url,
            "source_snapshots": list(source_snapshots.values()),
            "target_snapshots": list(target_snapshots.values()),
            "comparisons": comparisons,
            "identical_flows": identical_flows,
            "non_identical_flows": non_identical_flows,
            "missing_in_target_count": len(missing_in_target),
            "missing_in_target": missing_in_target,
            "error_count": len(errors),
            "errors": errors,
            "source_count": len(source_snapshots),
            "target_count": len(target_snapshots)
        }
