import pandas as pd

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')

# Convert to 1-based indexing
master_copy_df.index = master_copy_df.index + 1

# Create a mapping of finish codes to full names
finish_code_to_name = {}
for col in finishes_df.columns:
    finishes = finishes_df[col].dropna().tolist()
    for finish in finishes:
        if "(" in finish and ")" in finish:
            name = finish.split("(")[0].strip()
            code = finish.split("(")[1].split(")")[0].strip()
            finish_code_to_name[code] = name

# Get rows 14770-14774
rows = master_copy_df.loc[14770:14774].copy()

# Extract finish data
print("Finish data in rows 14770-14774:")
for idx, row in rows.iterrows():
    finish_value = row['finish'] if not pd.isna(row['finish']) else "NaN"
    sku = row['code'] if not pd.isna(row['code']) else "NaN"
    price = row['rrp'] if not pd.isna(row['rrp']) else "NaN"
    
    print(f"Row {idx}: Finish={finish_value}, SKU={sku}, Price={price}")
    
    # If this is a specific finish (not ## or x##), find the matching full name
    if finish_value in finish_code_to_name:
        print(f"  Maps to: {finish_code_to_name[finish_value]}")
    elif finish_value == "##":
        print("  ## finishes:")
        for code in ["PN", "SN", "BZ", "AB", "SB", "DB", "BAB", "BZW", "BABW", "ABW", "DBW", "NBW", "SBW", "PBUL"]:
            if code in finish_code_to_name:
                print(f"    {code} -> {finish_code_to_name[code]}")
    elif finish_value == "x##":
        print("  x## finishes:")
        for code in ["PCOP", "SCOP", "BLN", "PEW", "MBL", "ASV", "RGP", "ACOP"]:
            if code in finish_code_to_name:
                print(f"    {code} -> {finish_code_to_name[code]}")

# Look at column 25 in Finishes tab, which contains all finishes
all_finishes = finishes_df[25].dropna().tolist()
print("\nAll finishes in column 25:")
for idx, finish in enumerate(all_finishes):
    print(f"{idx+1}. {finish}")

# Check how many finishes are in the ## and x## categories
hash_codes = ["PN", "SN", "BZ", "AB", "SB", "DB", "BAB", "BZW", "BABW", "ABW", "DBW", "NBW", "SBW", "PBUL"]
xhash_codes = ["PCOP", "SCOP", "BLN", "PEW", "MBL", "ASV", "RGP", "ACOP"]

hash_finishes = [finish for finish in all_finishes if any(code in finish for code in hash_codes)]
xhash_finishes = [finish for finish in all_finishes if any(code in finish for code in xhash_codes)]

print(f"\nNumber of ## finishes: {len(hash_finishes)}")
print(f"Number of x## finishes: {len(xhash_finishes)}")
print(f"Combined total: {len(hash_finishes) + len(xhash_finishes)}")
print(f"Total of all finishes: {len(all_finishes)}")

# Special codes in MASTER COPY rows
special_codes = ["SCP", "PB", "##", "x##"]
print("\nMapping of special codes in MASTER COPY to finishes:")
for code in special_codes:
    if code in finish_code_to_name:
        print(f"{code} -> {finish_code_to_name[code]}")
    elif code == "##":
        print(f"{code} -> {len(hash_codes)} finishes: {hash_codes}")
    elif code == "x##":
        print(f"{code} -> {len(xhash_codes)} finishes: {xhash_codes}")
    else:
        print(f"{code} -> Unknown") 