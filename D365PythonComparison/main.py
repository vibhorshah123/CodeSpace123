"""
D365 Schema Comparison Tool
Main entry point for comparing Dynamics 365 table schemas between environments.
"""

import sys
import getpass
from typing import Dict, Any
from src.auth_manager import AuthManager
from src.schema_comparison import SchemaComparison
from src.data_comparison import DataComparison
from src.excel_generator import ExcelGenerator


def print_banner():
    """Display welcome banner"""
    print("=" * 70)
    print(" " * 10 + "Dynamics 365 Environment Comparison Tool")
    print("=" * 70)
    print()


def get_credentials() -> Dict[str, str]:
    """
    Prompt user for authentication credentials
    
    Returns:
        Dictionary containing authentication details
    """
    print("Please provide authentication details:")
    print("-" * 70)
    
    auth_method = input("Auth method (1=Interactive Login, 2=Service Principal): ").strip()
    
    credentials = {}
    
    if auth_method == "1":
        credentials["auth_type"] = "browser_login"
        print("\n  Your browser will open for sign-in.")
        print("  This works with MFA-enabled accounts and managed devices.")
    elif auth_method == "2":
        credentials["auth_type"] = "client_credentials"
        credentials["client_id"] = input("Client ID: ").strip()
        credentials["client_secret"] = getpass.getpass("Client Secret: ")
        credentials["tenant_id"] = input("Tenant ID: ").strip()
    else:
        print("Invalid option selected!")
        sys.exit(1)
    
    print()
    return credentials


def get_environment_details() -> Dict[str, str]:
    """
    Prompt user for environment URLs
    
    Returns:
        Dictionary containing environment URLs
    """
    print("Enter environment details:")
    print("-" * 70)
    
    env1_url = input("Source Environment URL (e.g., https://orgname.crm.dynamics.com): ").strip()
    env2_url = input("Target Environment URL (e.g., https://orgname2.crm.dynamics.com): ").strip()
    
    print()
    return {
        "source_url": env1_url,
        "target_url": env2_url
    }


def display_menu() -> str:
    """
    Display comparison options menu
    
    Returns:
        User's menu selection
    """
    print("\nSelect comparison type:")
    print("-" * 70)
    print("1. Schema Comparison (Compare table structures)")
    print("2. Data Comparison (Compare records + relationships)")
    print("3. Flow Comparison (Coming soon)")
    print("0. Exit")
    print("-" * 70)
    
    choice = input("Enter your choice: ").strip()
    return choice


def run_schema_comparison(auth_manager: AuthManager, envs: Dict[str, str]):
    """
    Execute schema comparison between two environments
    
    Args:
        auth_manager: Authentication manager instance
        envs: Dictionary containing source and target environment URLs
    """
    print("\n" + "=" * 70)
    print(" " * 20 + "Schema Comparison")
    print("=" * 70)
    
    # Prompt for table name
    table_name = input("\nEnter table logical name (e.g., contact, account, mash_customtable): ").strip()
    
    if not table_name:
        print("Error: Table name cannot be empty!")
        return
    
    print(f"\nFetching schema for table '{table_name}' from both environments...")
    
    try:
        # Initialize schema comparison
        comparison = SchemaComparison(auth_manager)
        
        # Perform comparison
        result = comparison.compare_table_schema(
            envs["source_url"],
            envs["target_url"],
            table_name
        )
        
        # Display summary
        print("\n" + "-" * 70)
        print("COMPARISON SUMMARY")
        print("-" * 70)
        print(f"Table: {table_name}")
        print(f"Source Environment: {envs['source_url']}")
        print(f"Target Environment: {envs['target_url']}")
        print()
        print(f"Fields only in Source: {len(result['only_in_source'])}")
        print(f"Fields only in Target: {len(result['only_in_target'])}")
        print(f"Fields with differences: {len(result['field_differences'])}")
        print(f"Matching fields: {len(result['matching_fields'])}")
        print("-" * 70)
        
        # Display details
        if result['only_in_source']:
            print("\nFields ONLY in Source:")
            for field in result['only_in_source']:
                print(f"  - {field}")
        
        if result['only_in_target']:
            print("\nFields ONLY in Target:")
            for field in result['only_in_target']:
                print(f"  - {field}")
        
        if result['field_differences']:
            print("\nFields with DIFFERENCES:")
            for diff in result['field_differences']:
                print(f"  - {diff['field_name']}:")
                for key, values in diff['differences'].items():
                    print(f"      {key}: Source='{values['source']}' | Target='{values['target']}'")
        
        # Generate Excel
        print("\n" + "-" * 70)
        generate_excel = input("Generate Excel report? (y/n): ").strip().lower()
        
        if generate_excel == 'y':
            output_file = input("Enter output filename (default: schema_comparison.xlsx): ").strip()
            if not output_file:
                output_file = "schema_comparison.xlsx"
            
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            
            generator = ExcelGenerator()
            generator.generate_schema_comparison_report(result, table_name, output_file)
            print(f"\n✓ Excel report generated: {output_file}")
        
    except Exception as e:
        print(f"\nError during schema comparison: {str(e)}")
        import traceback
        traceback.print_exc()


def run_data_comparison(auth_manager: AuthManager, envs: Dict[str, str]):
    """
    Execute data comparison between two environments
    
    Args:
        auth_manager: Authentication manager instance
        envs: Dictionary containing source and target environment URLs
    """
    print("\n" + "=" * 70)
    print(" " * 20 + "Data Comparison")
    print("=" * 70)
    
    # Prompt for table name
    table_name = input("\nEnter table logical name (e.g., contact, account, mash_customtable): ").strip()
    
    if not table_name:
        print("Error: Table name cannot be empty!")
        return
    
    # Ask about relationship comparison
    print("\nComparison Options:")
    print("  1. Compare main records only")
    print("  2. Compare main records + related records (One-To-Many relationships/subgrids)")
    
    choice = input("\nSelect option [1-2]: ").strip()
    include_relationships = choice == "2"
    
    print(f"\nFetching data for table '{table_name}' from both environments...")
    if include_relationships:
        print("  (Including related records via One-To-Many relationships)")
    
    try:
        # Initialize data comparison
        comparison = DataComparison(auth_manager)
        
        # Perform comparison
        result = comparison.compare_table_data(
            envs["source_url"],
            envs["target_url"],
            table_name,
            fields=None,  # Compare all fields
            include_relationships=include_relationships
        )
        
        # Display summary
        print("\n" + "-" * 70)
        print("COMPARISON SUMMARY")
        print("-" * 70)
        print(f"Table: {table_name}")
        print(f"Source Environment: {envs['source_url']}")
        print(f"Target Environment: {envs['target_url']}")
        print()
        print(f"Source records: {result['source_record_count']}")
        print(f"Target records: {result['target_record_count']}")
        print(f"Matching records: {len(result['matching_records'])}")
        print(f"Records only in Source: {len(result['only_in_source'])}")
        print(f"Records only in Target: {len(result['only_in_target'])}")
        print(f"Field mismatches: {len(result['mismatches'])}")
        
        if result.get('guid_mismatches'):
            print(f"\n   ⚠ GUID/Lookup Mismatches: {len(result['guid_mismatches'])}")
        
        if result.get('name_matches_with_different_ids'):
            print(f"   ⚠ Records with same name but different IDs: {len(result['name_matches_with_different_ids'])}")
            print(f"      (Possible duplicates or migration issues)")
        
        if include_relationships and result['child_comparisons']:
            print(f"\nRelated Entities Compared: {len(result['child_comparisons'])}")
            for child_entity, child_data in result['child_comparisons'].items():
                print(f"  - {child_entity}: {child_data['source_total']} source, {child_data['target_total']} target")
                if child_data['differences']:
                    print(f"      Differences found in {len(child_data['differences'])} parent records")
        
        print("-" * 70)
        
        # Generate Excel
        print("\n" + "-" * 70)
        generate_excel = input("Generate Excel report? (y/n): ").strip().lower()
        
        if generate_excel == 'y':
            output_file = input("Enter output filename (default: data_comparison.xlsx): ").strip()
            if not output_file:
                output_file = "data_comparison.xlsx"
            
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            
            generator = ExcelGenerator()
            generator.generate_data_comparison_report(result, table_name, output_file)
            print(f"\n✓ Excel report generated: {output_file}")
        
    except Exception as e:
        print(f"\nError during data comparison: {str(e)}")
        import traceback
        traceback.print_exc()


def main():
    """Main application entry point"""
    print_banner()
    
    # Get credentials
    credentials = get_credentials()
    
    # Get environment details
    envs = get_environment_details()
    
    # Initialize authentication manager
    print("Authenticating...")
    try:
        auth_manager = AuthManager(credentials)
        print("✓ Authentication successful!\n")
    except Exception as e:
        print(f"✗ Authentication failed: {str(e)}")
        sys.exit(1)
    
    # Main menu loop
    while True:
        choice = display_menu()
        
        if choice == "0":
            print("\nExiting... Goodbye!")
            break
        elif choice == "1":
            run_schema_comparison(auth_manager, envs)
        elif choice == "2":
            run_data_comparison(auth_manager, envs)
        elif choice == "3":
            print("\nFlow Comparison - Coming Soon!")
        else:
            print("\nInvalid choice! Please try again.")
        
        print("\n")
        input("Press Enter to continue...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user. Exiting...")
        sys.exit(0)
    except Exception as e:
        print(f"\n\nUnexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
