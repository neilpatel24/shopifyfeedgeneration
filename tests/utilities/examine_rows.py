import pandas as pd

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

# Read the specific rows we're interested in
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
rows = master_copy_df.loc[14770:14774].copy()

print("Detailed view of rows 14770-14774:")
print(rows[['description', 'size', 'code', 'rrp', 'finish', 'finish count']])

# Check for duplicate sizes
print("\nUnique sizes:")
unique_sizes = rows['size'].unique()
print(unique_sizes)
print(f"Number of unique sizes: {len(unique_sizes)}")

# Check for duplicate SKUs
print("\nUnique SKUs:")
unique_skus = rows['code'].unique()
print(unique_skus)
print(f"Number of unique SKUs: {len(unique_skus)}")

# Check for pricing variations
print("\nPricing by row:")
for idx, row in rows.iterrows():
    print(f"Row {idx}: Size={row['size']}, SKU={row['code']}, Price={row['rrp']}")

# Check the finishes column to determine number of finishes
print("\nFinish information:")
for idx, row in rows.iterrows():
    print(f"Row {idx}: Finish={row['finish']}, Finish Count={row['finish count']}")

# Check finishes tab
finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')
print("\nFinishes tab columns:")
print(finishes_df.columns.tolist())

# Check column 25 of Finishes tab
if 25 in finishes_df.columns:
    print("\nFinishes in column '25':")
    finishes = finishes_df[25].dropna().tolist()
    print(finishes)
    print(f"Number of finishes: {len(finishes)}")
elif '25' in finishes_df.columns:
    print("\nFinishes in column '25':")
    finishes = finishes_df['25'].dropna().tolist()
    print(finishes)
    print(f"Number of finishes: {len(finishes)}") 