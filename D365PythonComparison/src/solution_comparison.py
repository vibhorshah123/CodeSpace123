"""
Solution Comparison Module
Compares Dynamics 365 solution components between two environments
"""

import requests
from typing import Dict, List, Any, Optional
from src.auth_manager import AuthManager


class SolutionComparison:
    """Handles solution comparison operations between D365 environments"""
    
    def __init__(self, auth_manager: AuthManager):
        """
        Initialize solution comparison
        
        Args:
            auth_manager: Authentication manager instance
        """
        self.auth_manager = auth_manager
    
    def fetch_solutions(self, base_url: str, solution_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Fetch solutions from environment
        
        Args:
            base_url: Base URL for the environment
            solution_name: Optional solution unique name to filter by
            
        Returns:
            List of solution dictionaries
        """
        token = self.auth_manager.get_token(base_url)
        
        # Normalize URL
        if not base_url.startswith("https://"):
            base_url = f"https://{base_url}"
        base_url = base_url.rstrip('/')
        
        # Build API endpoint
        api_url = f"{base_url}/api/data/v9.2/solutions?$select=solutionid,uniquename,friendlyname,version,publisherid,ismanaged,installedon&$filter=isvisible eq true"
        
        if solution_name:
            # Escape single quotes in name
            escaped_name = solution_name.replace("'", "''")
            api_url = f"{base_url}/api/data/v9.2/solutions?$select=solutionid,uniquename,friendlyname,version,publisherid,ismanaged,installedon&$filter=isvisible eq true and uniquename eq '{escaped_name}'"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
        }
        
        try:
            response = requests.get(api_url, headers=headers, timeout=60)
            response.raise_for_status()
            data = response.json()
            
            if "value" in data:
                return data["value"]
            elif "solutionid" in data:
                return [data]
            return []
        except requests.exceptions.RequestException as e:
            print(f"    Error fetching solutions: {str(e)}")
            raise
    
    def fetch_solution_components(self, base_url: str, solution_id: str) -> List[Dict[str, Any]]:
        """
        Fetch components for a specific solution
        
        Args:
            base_url: Base URL for the environment
            solution_id: Solution ID (GUID)
            
        Returns:
            List of solution component dictionaries
        """
        token = self.auth_manager.get_token(base_url)
        
        # Normalize URL
        if not base_url.startswith("https://"):
            base_url = f"https://{base_url}"
        base_url = base_url.rstrip('/')
        
        # Build API endpoint for solution components
        api_url = f"{base_url}/api/data/v9.2/solutioncomponents?$select=solutioncomponentid,componenttype,objectid,rootcomponentbehavior&$filter=_solutionid_value eq {solution_id}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0"
        }
        
        components = []
        
        print(f"  Fetching components for solution {solution_id}...")
        
        while api_url:
            try:
                response = requests.get(api_url, headers=headers, timeout=60)
                response.raise_for_status()
                data = response.json()
                
                if "value" in data:
                    components.extend(data["value"])
                
                # Check for next page
                api_url = data.get("@odata.nextLink")
                
            except requests.exceptions.RequestException as e:
                print(f"    Error fetching components: {str(e)}")
                raise
        
        print(f"  ✓ Total components fetched: {len(components)}")
        return components
    
    def get_component_type_name(self, component_type: int) -> str:
        """
        Get human-readable component type name
        
        Args:
            component_type: Component type code
            
        Returns:
            Component type name
        """
        component_types = {
            1: "Entity",
            2: "Attribute",
            3: "Relationship",
            4: "Attribute Picklist Value",
            5: "Attribute Lookup Value",
            9: "Option Set",
            10: "Entity Relationship",
            11: "Entity Relationship Role",
            12: "Entity Relationship Relationships",
            13: "Managed Property",
            14: "Entity Key",
            16: "Privilege",
            17: "PrivilegeObjectTypeCode",
            20: "Role",
            21: "Role Privilege",
            22: "Display String",
            23: "Display String Map",
            24: "Form",
            25: "Organization",
            26: "Saved Query",
            29: "Workflow",
            31: "Report",
            32: "Report Entity",
            33: "Report Category",
            34: "Report Visibility",
            35: "Attachment",
            36: "Email Template",
            37: "Contract Template",
            38: "KB Article Template",
            39: "Mail Merge Template",
            44: "Duplicate Rule",
            45: "Duplicate Rule Condition",
            46: "Entity Map",
            47: "Attribute Map",
            48: "Ribbon Command",
            49: "Ribbon Context Group",
            50: "Ribbon Customization",
            52: "Ribbon Rule",
            53: "Ribbon Tab To Command Map",
            55: "Ribbon Diff",
            59: "Saved Query Visualization",
            60: "System Form",
            61: "Web Resource",
            62: "Site Map",
            63: "Connection Role",
            64: "Complex Control",
            65: "Hierarchy Rule",
            66: "Custom Control",
            68: "Custom Control Default Config",
            70: "Field Security Profile",
            71: "Field Permission",
            80: "Plugin Type",
            81: "Plugin Assembly",
            82: "SDK Message Processing Step",
            83: "SDK Message Processing Step Image",
            84: "Service Endpoint",
            90: "Report",
            91: "Report Related Entity",
            92: "Report Related Category",
            93: "Report Visibility",
            95: "App Module Metadata",
            96: "App Module Component",
            152: "Mobile Offline Profile",
            153: "Mobile Offline Profile Item",
            154: "Mobile Offline Profile Item Association",
            161: "Connector",
            162: "Environment Variable Definition",
            163: "Environment Variable Value",
            165: "AI Project Type",
            166: "AI Project",
            167: "AI Configuration",
            168: "Entity Analytics Configuration",
            300: "Canvas App"
        }
        return component_types.get(component_type, f"Unknown ({component_type})")
    
    def compare_solutions(self, 
                         source_url: str,
                         target_url: str,
                         solution_name: str) -> Dict[str, Any]:
        """
        Compare a solution between two environments
        
        Args:
            source_url: Source environment URL
            target_url: Target environment URL
            solution_name: Solution unique name to compare
            
        Returns:
            Dictionary containing comparison results
        """
        print(f"\nFetching solution '{solution_name}' from source environment...")
        source_solutions = self.fetch_solutions(source_url, solution_name)
        
        if not source_solutions:
            raise Exception(f"Solution '{solution_name}' not found in source environment")
        
        source_solution = source_solutions[0]
        
        print(f"\nFetching solution '{solution_name}' from target environment...")
        target_solutions = self.fetch_solutions(target_url, solution_name)
        
        if not target_solutions:
            print(f"  ⚠ Solution '{solution_name}' not found in target environment")
            return {
                "source_url": source_url,
                "target_url": target_url,
                "solution_name": solution_name,
                "source_solution": source_solution,
                "target_solution": None,
                "source_components": [],
                "target_components": [],
                "status": "missing_in_target",
                "component_summary": {},
                "only_in_source": [],
                "only_in_target": [],
                "common_components": []
            }
        
        target_solution = target_solutions[0]
        
        # Fetch components
        print(f"\nFetching components from source environment...")
        source_components = self.fetch_solution_components(source_url, source_solution['solutionid'])
        
        print(f"\nFetching components from target environment...")
        target_components = self.fetch_solution_components(target_url, target_solution['solutionid'])
        
        # Compare components
        print(f"\nComparing solution components...")
        
        # Create component maps by objectid and componenttype
        source_comp_map = {
            (comp['objectid'], comp['componenttype']): comp 
            for comp in source_components
        }
        target_comp_map = {
            (comp['objectid'], comp['componenttype']): comp 
            for comp in target_components
        }
        
        source_keys = set(source_comp_map.keys())
        target_keys = set(target_comp_map.keys())
        
        only_in_source = []
        only_in_target = []
        common_components = []
        
        for key in source_keys - target_keys:
            comp = source_comp_map[key]
            only_in_source.append({
                "objectid": comp['objectid'],
                "componenttype": comp['componenttype'],
                "componenttype_name": self.get_component_type_name(comp['componenttype']),
                "rootcomponentbehavior": comp.get('rootcomponentbehavior')
            })
        
        for key in target_keys - source_keys:
            comp = target_comp_map[key]
            only_in_target.append({
                "objectid": comp['objectid'],
                "componenttype": comp['componenttype'],
                "componenttype_name": self.get_component_type_name(comp['componenttype']),
                "rootcomponentbehavior": comp.get('rootcomponentbehavior')
            })
        
        for key in source_keys & target_keys:
            comp = source_comp_map[key]
            common_components.append({
                "objectid": comp['objectid'],
                "componenttype": comp['componenttype'],
                "componenttype_name": self.get_component_type_name(comp['componenttype']),
                "rootcomponentbehavior": comp.get('rootcomponentbehavior')
            })
        
        # Group by component type for summary
        component_summary = {}
        
        for comp in source_components:
            comp_type = comp['componenttype']
            type_name = self.get_component_type_name(comp_type)
            if type_name not in component_summary:
                component_summary[type_name] = {"source": 0, "target": 0, "common": 0, "source_only": 0, "target_only": 0}
            component_summary[type_name]["source"] += 1
        
        for comp in target_components:
            comp_type = comp['componenttype']
            type_name = self.get_component_type_name(comp_type)
            if type_name not in component_summary:
                component_summary[type_name] = {"source": 0, "target": 0, "common": 0, "source_only": 0, "target_only": 0}
            component_summary[type_name]["target"] += 1
        
        for comp in common_components:
            type_name = comp['componenttype_name']
            if type_name in component_summary:
                component_summary[type_name]["common"] += 1
        
        for comp in only_in_source:
            type_name = comp['componenttype_name']
            if type_name in component_summary:
                component_summary[type_name]["source_only"] += 1
        
        for comp in only_in_target:
            type_name = comp['componenttype_name']
            if type_name in component_summary:
                component_summary[type_name]["target_only"] += 1
        
        print(f"  ✓ Comparison complete")
        
        return {
            "source_url": source_url,
            "target_url": target_url,
            "solution_name": solution_name,
            "source_solution": source_solution,
            "target_solution": target_solution,
            "source_components": source_components,
            "target_components": target_components,
            "status": "compared",
            "component_summary": component_summary,
            "only_in_source": only_in_source,
            "only_in_target": only_in_target,
            "common_components": common_components,
            "source_component_count": len(source_components),
            "target_component_count": len(target_components),
            "common_count": len(common_components),
            "source_only_count": len(only_in_source),
            "target_only_count": len(only_in_target)
        }
