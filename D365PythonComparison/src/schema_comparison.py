"""
Schema Comparison Module
Compares table schemas between two Dynamics 365 environments
"""

import requests
from typing import Dict, List, Any
from src.auth_manager import AuthManager


class SchemaComparison:
    """Handles schema comparison operations between D365 environments"""
    
    def __init__(self, auth_manager: AuthManager):
        """
        Initialize schema comparison
        
        Args:
            auth_manager: AuthManager instance for authentication
        """
        self.auth_manager = auth_manager
    
    def get_table_metadata(self, environment_url: str, table_logical_name: str) -> Dict[str, Any]:
        """
        Fetch table metadata from D365 environment
        
        Args:
            environment_url: URL of the D365 environment
            table_logical_name: Logical name of the table (e.g., 'contact', 'account')
            
        Returns:
            Dictionary containing table metadata
        """
        token = self.auth_manager.get_token(environment_url)
        
        # Construct API URL
        base_url = environment_url.rstrip('/')
        api_url = f"{base_url}/api/data/v9.2/EntityDefinitions(LogicalName='{table_logical_name}')"
        
        # Build query string manually - only include properties available on all AttributeMetadata types
        # Note: MaxLength, Precision are type-specific properties, so we fetch full attributes and filter later
        query_params = (
            "$select=LogicalName,DisplayName,SchemaName,ObjectTypeCode,PrimaryIdAttribute,PrimaryNameAttribute"
            "&$expand=Attributes($select=LogicalName,AttributeType,DisplayName,SchemaName,RequiredLevel,"
            "IsValidForCreate,IsValidForUpdate,IsValidForRead,IsPrimaryId,IsPrimaryName)"
        )
        
        api_url = f"{api_url}?{query_params}"
        
        params = None  # Using full URL instead of params dict
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Content-Type": "application/json"
        }
        
        try:
            response = requests.get(api_url, headers=headers, timeout=60)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            # Get more details from response if available
            error_detail = ""
            try:
                if hasattr(e, 'response') and e.response is not None:
                    error_detail = f"\nResponse: {e.response.text[:500]}"
            except:
                pass
            raise Exception(f"Failed to fetch metadata from {environment_url}: {str(e)}{error_detail}")
    
    def normalize_attribute(self, attr: Dict[str, Any]) -> Dict[str, Any]:
        """
        Normalize attribute metadata for comparison
        
        Args:
            attr: Attribute metadata dictionary
            
        Returns:
            Normalized attribute dictionary
        """
        # Helper function to safely extract display name
        def get_display_name(display_name_obj):
            if display_name_obj is None:
                return ""
            if isinstance(display_name_obj, dict):
                user_label = display_name_obj.get("UserLocalizedLabel")
                if user_label and isinstance(user_label, dict):
                    return user_label.get("Label", "")
                return ""
            return str(display_name_obj)
        
        # Helper function to safely extract required level
        def get_required_level(required_level_obj):
            if required_level_obj is None:
                return ""
            if isinstance(required_level_obj, dict):
                return required_level_obj.get("Value", "")
            return str(required_level_obj)
        
        # Extract key properties for comparison
        normalized = {
            "LogicalName": attr.get("LogicalName", ""),
            "AttributeType": attr.get("AttributeType", ""),
            "SchemaName": attr.get("SchemaName", ""),
            "DisplayName": get_display_name(attr.get("DisplayName")),
            "RequiredLevel": get_required_level(attr.get("RequiredLevel")),
            "IsValidForCreate": attr.get("IsValidForCreate", False),
            "IsValidForUpdate": attr.get("IsValidForUpdate", False),
            "IsValidForRead": attr.get("IsValidForRead", False),
            "IsPrimaryId": attr.get("IsPrimaryId", False),
            "IsPrimaryName": attr.get("IsPrimaryName", False)
        }
        
        # Note: Type-specific properties like MaxLength and Precision are not included
        # in the initial API query to avoid errors. If needed in the future, we can
        # make additional type-specific API calls for those attributes.
        
        return normalized
    
    def compare_attributes(self, source_attr: Dict[str, Any], target_attr: Dict[str, Any]) -> Dict[str, Any]:
        """
        Compare two attributes and identify differences
        
        Args:
            source_attr: Normalized source attribute
            target_attr: Normalized target attribute
            
        Returns:
            Dictionary of differences (empty if identical)
        """
        differences = {}
        
        # Compare each property
        for key in source_attr.keys():
            source_val = source_attr.get(key)
            target_val = target_attr.get(key)
            
            if source_val != target_val:
                differences[key] = {
                    "source": source_val,
                    "target": target_val
                }
        
        return differences
    
    def compare_table_schema(self, source_url: str, target_url: str, table_logical_name: str) -> Dict[str, Any]:
        """
        Compare table schema between two environments
        
        Args:
            source_url: Source environment URL
            target_url: Target environment URL
            table_logical_name: Logical name of the table to compare
            
        Returns:
            Dictionary containing comparison results
        """
        print(f"  Fetching source metadata from {source_url}...")
        source_metadata = self.get_table_metadata(source_url, table_logical_name)
        
        print(f"  Fetching target metadata from {target_url}...")
        target_metadata = self.get_table_metadata(target_url, table_logical_name)
        
        # Extract and normalize attributes
        source_attributes = source_metadata.get("Attributes", [])
        target_attributes = target_metadata.get("Attributes", [])
        
        print(f"  Analyzing {len(source_attributes)} source fields and {len(target_attributes)} target fields...")
        
        # Create dictionaries for easy lookup
        source_attr_dict = {
            attr["LogicalName"]: self.normalize_attribute(attr)
            for attr in source_attributes
        }
        
        target_attr_dict = {
            attr["LogicalName"]: self.normalize_attribute(attr)
            for attr in target_attributes
        }
        
        # Find differences
        source_fields = set(source_attr_dict.keys())
        target_fields = set(target_attr_dict.keys())
        
        only_in_source = sorted(list(source_fields - target_fields))
        only_in_target = sorted(list(target_fields - source_fields))
        common_fields = source_fields & target_fields
        
        # Compare common fields
        field_differences = []
        matching_fields = []
        
        for field_name in sorted(common_fields):
            source_attr = source_attr_dict[field_name]
            target_attr = target_attr_dict[field_name]
            
            diff = self.compare_attributes(source_attr, target_attr)
            
            if diff:
                field_differences.append({
                    "field_name": field_name,
                    "differences": diff,
                    "source_data": source_attr,
                    "target_data": target_attr
                })
            else:
                matching_fields.append(field_name)
        
        return {
            "table_name": table_logical_name,
            "source_url": source_url,
            "target_url": target_url,
            "only_in_source": only_in_source,
            "only_in_target": only_in_target,
            "field_differences": field_differences,
            "matching_fields": matching_fields,
            "source_metadata": source_metadata,
            "target_metadata": target_metadata
        }
