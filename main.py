import pandas as pd
from constant import TYPE_ALIASES
import datetime

# Get the current timestamp
current_time = datetime.datetime.now().strftime("%H%M%S")

SHOPEE = {"url": "files/DAFTAR PRODUK SHOPEE.xlsx", "skiprows": 1, "skipcols": 2}
TIKTOK = {"url": "files/DAFTAR PRODUK TIKTOK.xlsx", "skiprows": 5, "skipcols": 3}
LAZADA = {"url": "files/DAFTAR PRODUK LAZADA.xlsx", "skiprows": 4, "skipcols": 2}
TOKPED = {"url": "files/DAFTAR PRODUK TOKPED.xlsx", "skiprows": 3, "skipcols": 2}


def find_items_by_product_types(product_names, type_aliases):
    results = {main_type: [] for main_type in type_aliases}
    remaining_items = []

    for product_name in product_names:
        found = False

        for main_type, aliases in type_aliases.items():
            # Check if any alias is found in the product name
            for alias in aliases:
                if alias in product_name.lower():
                    results[main_type].append(product_name)
                    found = True
                    break  # Exit if a match is found for this alias
            if found:
                break  # Exit if a match is found for this main type

        if not found:
            remaining_items.append(product_name)

    # Sort results
    for main_type in results:
        results[main_type] = sorted(results[main_type])

    return results, sorted(remaining_items)


# Example usage with your Excel data

# Load the entire sheet, starting from row 6
df = pd.read_excel(LAZADA["url"], skiprows=LAZADA["skiprows"])
column_d_data = df.iloc[:, LAZADA["skipcols"]].dropna()

# Call the function with cleaned data and the alias dictionary
results, remaining_items = find_items_by_product_types(column_d_data, TYPE_ALIASES)

# Write matched and unmatched items to separate files
with open(f"output/output_LAZADA_{current_time}_matched.txt", "w") as matched_file:
    for type_name, items in results.items():
        if items:
            matched_file.write(f"\n{type_name.capitalize()} Items ({len(items)}):\n")
            for item in items:
                matched_file.write(f"  - {item}\n")

if remaining_items:
    with open(
        f"output/output_LAZADA_{current_time}_unmatched.txt", "w"
    ) as unmatched_file:
        unmatched_file.write(f"Remaining Unmatched Items ({len(remaining_items)}):\n")
        for item in remaining_items:
            unmatched_file.write(f"  - {item}\n")
