"""
D365 Schema Comparison Tool
Main entry point for comparing Dynamics 365 table schemas between environments.
"""

import sys
import getpass
from typing import Dict, Any
from datetime import datetime
from src.auth_manager import AuthManager
from src.schema_comparison import SchemaComparison
from src.data_comparison import DataComparison
from src.flow_comparison import FlowComparison
from src.solution_comparison import SolutionComparison
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
    Prompt user for environment names and construct URLs
    
    Returns:
        Dictionary containing environment URLs
    """
    print("Enter environment details:")
    print("-" * 70)
    
    env1_name = input("Source Environment Name (e.g., mash, orgname): ").strip()
    env2_name = input("Target Environment Name (e.g., mashvnext, orgname2): ").strip()
    
    # Construct full URLs
    env1_url = f"https://{env1_name}.crm.dynamics.com"
    env2_url = f"https://{env2_name}.crm.dynamics.com"
    
    print(f"\nSource URL: {env1_url}")
    print(f"Target URL: {env2_url}")
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
    print("3. Flow Comparison (Compare Power Automate flows)")
    print("4. Solution Comparison (Compare solution components)")
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
            # Auto-generate filename: schema_comparison_{entityname}_{datenow}.xlsx
            date_now = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"schema_comparison_{table_name}_{date_now}.xlsx"
            
            # Extract environment names from URLs
            source_env_name = envs["source_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            target_env_name = envs["target_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            
            generator = ExcelGenerator()
            generator.generate_schema_comparison_report(result, table_name, output_file, source_env_name, target_env_name)
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
    
    print(f"\nFetching data for table '{table_name}' from both environments...")
    print(f"Match Mode: PRIMARY NAME")
    
    try:
        # Initialize data comparison
        comparison = DataComparison(auth_manager)
        
        # Perform comparison
        result = comparison.compare_table_data(
            envs["source_url"],
            envs["target_url"],
            table_name,
            fields=None  # Compare all fields
        )
        
        # Display summary
        print("\n" + "-" * 70)
        print("COMPARISON SUMMARY")
        print("-" * 70)
        print(f"Table: {table_name}")
        print(f"Source Environment: {envs['source_url']}")
        print(f"Target Environment: {envs['target_url']}")
        print(f"Match Mode: PRIMARY NAME")
        print()
        print(f"Source records: {result['source_record_count']}")
        print(f"Target records: {result['target_record_count']}")
        print(f"Matching records: {len(result['matching_records'])}")
        print(f"Records only in Source: {len(result['only_in_source'])}")
        print(f"Records only in Target: {len(result['only_in_target'])}")
        print(f"Field mismatches: {len(result['mismatches'])}")
        
        # Name-based info
        if result.get('name_matched_records'):
            print(f"\n   ℹ Records matched by name (different GUIDs observed): {len(result['name_matched_records'])}")
        if result.get('duplicate_names_source'):
            print(f"   ⚠ Duplicate names in source: {len(result['duplicate_names_source'])}")
        if result.get('duplicate_names_target'):
            print(f"   ⚠ Duplicate names in target: {len(result['duplicate_names_target'])}")
        
        print("-" * 70)
        
        # Generate Excel
        print("\n" + "-" * 70)
        generate_excel = input("Generate Excel report? (y/n): ").strip().lower()
        
        if generate_excel == 'y':
            # Auto-generate filename: data_comparison_{entityname}_{datenow}.xlsx
            date_now = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"data_comparison_{table_name}_{date_now}.xlsx"
            
            # Extract environment names from URLs
            source_env_name = envs["source_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            target_env_name = envs["target_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            
            generator = ExcelGenerator()
            generator.generate_data_comparison_report(result, table_name, output_file, source_env_name, target_env_name)
            print(f"\n✓ Excel report generated: {output_file}")
        
    except Exception as e:
        print(f"\nError during data comparison: {str(e)}")
        import traceback
        traceback.print_exc()


def run_flow_comparison(auth_manager: AuthManager, envs: Dict[str, str]):
    """
    Execute flow comparison between two environments
    
    Args:
        auth_manager: Authentication manager instance
        envs: Dictionary containing source and target environment URLs
    """
    print("\n" + "=" * 70)
    print(" " * 20 + "Flow Comparison")
    print("=" * 70)
    
    # Prompt for flow name
    flow_name = input("\nEnter flow name to compare: ").strip()
    
    if not flow_name:
        print("Error: Flow name cannot be empty!")
        return
    
    print(f"\nComparing flow '{flow_name}' between environments...")
    
    try:
        # Initialize flow comparison
        comparison = FlowComparison(auth_manager)
        
        # Perform comparison
        result = comparison.compare_flows(
            envs["source_url"],
            envs["target_url"],
            flow_name=flow_name,
            include_diff_details=True
        )
        
        # Display summary
        print("\n" + "-" * 70)
        print("COMPARISON SUMMARY")
        print("-" * 70)
        print(f"Source Environment: {envs['source_url']}")
        print(f"Target Environment: {envs['target_url']}")
        print()
        print(f"Source flows: {result['source_count']}")
        print(f"Target flows: {result['target_count']}")
        print(f"Identical flows: {len(result['identical_flows'])}")
        print(f"Different flows: {len(result['non_identical_flows'])}")
        print(f"Missing in Target: {result['missing_in_target_count']}")
        print(f"Flows with errors: {result['error_count']}")
        print("-" * 70)
        
        # Show some details
        if result['identical_flows']:
            print(f"\nIdentical Flows ({len(result['identical_flows'])}):")
            for flow in result['identical_flows'][:5]:
                print(f"  ✓ {flow}")
            if len(result['identical_flows']) > 5:
                print(f"  ... and {len(result['identical_flows']) - 5} more")
        
        if result['non_identical_flows']:
            print(f"\nDifferent Flows ({len(result['non_identical_flows'])}):")
            for flow in result['non_identical_flows'][:5]:
                print(f"  ⚠ {flow}")
            if len(result['non_identical_flows']) > 5:
                print(f"  ... and {len(result['non_identical_flows']) - 5} more")
        
        if result['missing_in_target']:
            print(f"\nMissing in Target ({result['missing_in_target_count']}):")
            for flow in result['missing_in_target'][:5]:
                print(f"  ✗ {flow}")
            if result['missing_in_target_count'] > 5:
                print(f"  ... and {result['missing_in_target_count'] - 5} more")
        
        # Generate Excel
        print("\n" + "-" * 70)
        generate_excel = input("Generate Excel report? (y/n): ").strip().lower()
        
        if generate_excel == 'y':
            # Auto-generate filename: flow_comparison_{flowname}_{datenow}.xlsx
            date_now = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = flow_name.replace(" ", "_").replace("/", "_")
            output_file = f"flow_comparison_{safe_name}_{date_now}.xlsx"
            
            # Extract environment names from URLs
            source_env_name = envs["source_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            target_env_name = envs["target_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            
            generator = ExcelGenerator()
            generator.generate_flow_comparison_report(result, output_file, source_env_name, target_env_name)
            print(f"\n✓ Excel report generated: {output_file}")
        
    except Exception as e:
        print(f"\nError during flow comparison: {str(e)}")
        import traceback
        traceback.print_exc()


def run_solution_comparison(auth_manager: AuthManager, envs: Dict[str, str]):
    """
    Execute solution comparison between two environments
    
    Args:
        auth_manager: Authentication manager instance
        envs: Dictionary containing source and target environment URLs
    """
    print("\n" + "=" * 70)
    print(" " * 20 + "Solution Comparison")
    print("=" * 70)
    
    # Prompt for solution unique name
    solution_name = input("\nEnter solution unique name (e.g., mysolution): ").strip()
    
    if not solution_name:
        print("Error: Solution name cannot be empty!")
        return
    
    print(f"\nComparing solution '{solution_name}' between environments...")
    
    try:
        # Initialize solution comparison
        comparison = SolutionComparison(auth_manager)
        
        # Perform comparison
        result = comparison.compare_solutions(
            envs["source_url"],
            envs["target_url"],
            solution_name
        )
        
        # Check if solution was found
        if result['status'] == 'missing_in_target':
            print(f"\n⚠ Warning: Solution '{solution_name}' found in source but NOT in target environment!")
            print(f"  Source components: {result['source_component_count']}")
            
            # Show component type summary
            if result.get('component_summary'):
                print("\n  Component breakdown:")
                for comp_type, stats in sorted(result['component_summary'].items()):
                    if stats['source'] > 0:
                        print(f"    {comp_type}: {stats['source']}")
            
            # Ask if user wants to generate report anyway
            print("\n" + "-" * 70)
            generate_excel = input("Generate Excel report? (y/n): ").strip().lower()
            
            if generate_excel == 'y':
                date_now = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"solution_comparison_{solution_name}_{date_now}.xlsx"
                
                source_env_name = envs["source_url"].replace("https://", "").replace(".crm.dynamics.com", "")
                target_env_name = envs["target_url"].replace("https://", "").replace(".crm.dynamics.com", "")
                
                generator = ExcelGenerator()
                generator.generate_solution_comparison_report(result, output_file, source_env_name, target_env_name)
                print(f"\n✓ Excel report generated: {output_file}")
            
            return
        
        # Display summary
        print("\n" + "-" * 70)
        print("COMPARISON SUMMARY")
        print("-" * 70)
        print(f"Solution: {solution_name}")
        print(f"Source Environment: {envs['source_url']}")
        print(f"Target Environment: {envs['target_url']}")
        print()
        
        # Solution details
        if result.get('source_solution'):
            source_sol = result['source_solution']
            print(f"Source Version: {source_sol.get('version', 'N/A')}")
            print(f"Source Is Managed: {'Yes' if source_sol.get('ismanaged') else 'No'}")
        
        if result.get('target_solution'):
            target_sol = result['target_solution']
            print(f"Target Version: {target_sol.get('version', 'N/A')}")
            print(f"Target Is Managed: {'Yes' if target_sol.get('ismanaged') else 'No'}")
        
        print()
        print(f"Source Components: {result['source_component_count']}")
        print(f"Target Components: {result['target_component_count']}")
        print(f"Common Components: {result['common_count']}")
        print(f"Only in Source: {result['source_only_count']}")
        print(f"Only in Target: {result['target_only_count']}")
        print("-" * 70)
        
        # Show component type summary
        if result.get('component_summary'):
            print("\nComponent Type Summary:")
            print("-" * 70)
            print(f"{'Component Type':<30} {'Source':<10} {'Target':<10} {'Common':<10}")
            print("-" * 70)
            
            for comp_type, stats in sorted(result['component_summary'].items()):
                print(f"{comp_type:<30} {stats['source']:<10} {stats['target']:<10} {stats['common']:<10}")
        
        # Show differences
        if result['source_only_count'] > 0:
            print(f"\nComponents Only in Source ({result['source_only_count']}):")
            only_in_source = result['only_in_source'][:10]
            for comp in only_in_source:
                print(f"  ⚠ {comp['componenttype_name']}: {comp['objectid']}")
            if result['source_only_count'] > 10:
                print(f"  ... and {result['source_only_count'] - 10} more")
        
        if result['target_only_count'] > 0:
            print(f"\nComponents Only in Target ({result['target_only_count']}):")
            only_in_target = result['only_in_target'][:10]
            for comp in only_in_target:
                print(f"  ℹ {comp['componenttype_name']}: {comp['objectid']}")
            if result['target_only_count'] > 10:
                print(f"  ... and {result['target_only_count'] - 10} more")
        
        # Generate Excel
        print("\n" + "-" * 70)
        generate_excel = input("Generate Excel report? (y/n): ").strip().lower()
        
        if generate_excel == 'y':
            date_now = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"solution_comparison_{solution_name}_{date_now}.xlsx"
            
            source_env_name = envs["source_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            target_env_name = envs["target_url"].replace("https://", "").replace(".crm.dynamics.com", "")
            
            generator = ExcelGenerator()
            generator.generate_solution_comparison_report(result, output_file, source_env_name, target_env_name)
            print(f"\n✓ Excel report generated: {output_file}")
        
    except Exception as e:
        print(f"\nError during solution comparison: {str(e)}")
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
            run_flow_comparison(auth_manager, envs)
        elif choice == "4":
            run_solution_comparison(auth_manager, envs)
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
