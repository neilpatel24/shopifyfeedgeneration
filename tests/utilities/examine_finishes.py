import pandas as pd

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'
finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')

# Display all columns
print("All columns in Finishes tab:")
print(finishes_df.columns.tolist())

# For each column, show the finishes
for col in finishes_df.columns:
    finishes = finishes_df[col].dropna().tolist()
    print(f"\nFinishes in column '{col}' ({len(finishes)} finishes):")
    for idx, finish in enumerate(finishes):
        # Try to extract finish code (typically in parentheses)
        finish_code = ""
        if "(" in finish and ")" in finish:
            finish_code = finish.split("(")[1].split(")")[0]
        print(f"{idx+1}. {finish} - Code: {finish_code}")

# Extract mapping of finish codes to full names
print("\n\nFinish code to name mapping:")
all_finishes = {}
for col in finishes_df.columns:
    finishes = finishes_df[col].dropna().tolist()
    for finish in finishes:
        if "(" in finish and ")" in finish:
            name = finish.split("(")[0].strip()
            code = finish.split("(")[1].split(")")[0].strip()
            all_finishes[code] = name
            print(f"{code} -> {name}")

# Check the specific finishes mentioned
print("\nChecking mentioned finish codes:")
##_finishes = ["PN", "SN", "BZ", "AB", "SB", "DB", "BAB", "BZW", "BABW", "ABW", "DBW", "NBW", "SBW", "PBUL"]
## _finishes_full = []
for code in ["PN", "SN", "BZ", "AB", "SB", "DB", "BAB", "BZW", "BABW", "ABW", "DBW", "NBW", "SBW", "PBUL"]:
    if code in all_finishes:
        print(f"{code} -> {all_finishes[code]}")
    else:
        print(f"{code} -> Not found")

print("\nChecking X## finish codes:")
for code in ["PCOP", "SCOP", "BLN", "PEW", "MBL", "ASV", "RGP", "ACOP"]:
    if code in all_finishes:
        print(f"{code} -> {all_finishes[code]}")
    else:
        print(f"{code} -> Not found") 