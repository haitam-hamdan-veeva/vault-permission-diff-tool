import json

import pandas as pd
import requests


def load_settings(file_path):
    with open(file_path) as config_file:
        return json.load(config_file)


config_data = load_settings("config.json")
vault_dns = config_data["vault_settings"]["vault_dns"]
api_version = config_data["vault_settings"]["api_version"]
session_id = config_data["vault_settings"]["session_id"]
source_security_profile_key = config_data["security_profiles_settings"][
    "source_security_profile_key"
]
target_security_profile_key = config_data["security_profiles_settings"][
    "target_security_profile_key"
]


def set_api_url(component_type, component_attribute):
    return f"https://{vault_dns}/api/{api_version}/configuration/{component_type}.{component_attribute}"


def make_api_request(url, headers):
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()


def get_permission_sets(security_profile_key):
    url = set_api_url("Securityprofile", security_profile_key)
    headers = {
        "Authorization": f"Bearer {session_id}",
        "Content-Type": "application/json",
    }
    data = make_api_request(url, headers)
    return data.get("data", {}).get("permission_sets", []) if data else []


def get_permissions(permission_set_key):
    url = set_api_url("Permissionset", permission_set_key)
    headers = {
        "Authorization": f"Bearer {session_id}",
        "Content-Type": "application/json",
    }
    data = make_api_request(url, headers)
    return data.get("data", {}).get("permission", []) if data else []


def process_permissions(security_profile_key):
    permission_sets = get_permission_sets(security_profile_key)
    all_permissions = []

    for permission_set in permission_sets:
        permissions = get_permissions(permission_set)

        for permission in permissions:
            permission_dict = {
                "Object": permission.get("object"),
                "Permission Group": permission.get("permission_group"),
                "Permission Subgroup": permission.get("permission_subgroup"),
                "Permission List": tuple(sorted(permission.get("permission_list", []))),
            }
            all_permissions.append(permission_dict)
    return all_permissions


def create_dataframe(permissions_data):
    return pd.DataFrame(permissions_data)


def compare_dataframes(dataframe_1, dataframe_2):
    comparison_dataframe = pd.merge(
        dataframe_1,
        dataframe_2,
        on=["Object", "Permission Group", "Permission Subgroup", "Permission List"],
        how="outer",
        indicator=True,
    )
    comparison_dataframe["Diff"] = comparison_dataframe["_merge"].map(
        {
            "left_only": f"Only in Source: {dataframe_1}",
            "right_only": f"Only in Target: {dataframe_2}",
            "both": "In Both",
        }
    )
    return comparison_dataframe


def highlight_differences(row):
    return [
        (
            "background-color: lightcoral"
            if row["Diff"] in ["Only in Source", "Only in Target"]
            else "background-color: lightgreen"
        )
    ] * len(row)


def save_to_excel(dataframe, file_path, sheet_name):
    dataframe.to_excel(file_path, sheet_name, index=False, engine="openpyxl")
    print(f"Permission Diff results were written successfully to {file_path}")


if __name__ == "__main__":
    # Process permissions for source and target profiles
    source_permissions = process_permissions(source_security_profile_key)
    target_permissions = process_permissions(target_security_profile_key)

    print(source_permissions)
    print(target_permissions)

    # Create DataFrames
    source_dataframe = create_dataframe(source_permissions)
    target_dataframe = create_dataframe(target_permissions)

    print(source_dataframe)
    print(target_dataframe)

    # Compare DataFrames
    comparison_dataframe = compare_dataframes(source_dataframe, target_dataframe)
    print(comparison_dataframe)

    # Remove duplicates
    unique_dataframe = comparison_dataframe.drop_duplicates(
        subset=["Object", "Permission Group", "Permission Subgroup", "Permission List"]
    )

    # Highlight differences
    styled_dataframe = unique_dataframe.style.apply(highlight_differences, axis=1)

    # Save results to Excel
    save_to_excel(
        styled_dataframe, "permissions_diff_results.xlsx", "Permission Diff Results"
    )
