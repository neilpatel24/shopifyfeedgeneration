import pandas as pd

# Load the output file
df = pd.read_excel('test_updated.xlsx')

# Print basic info
print(f"Total rows in output: {len(df)}")
print(f"Unique sizes: {df['Option1 Value'].unique()}")
print(f"Number of unique sizes: {len(df['Option1 Value'].unique())}")
print(f"Unique finishes: {len(df['Option2 Value'].unique())}")

# Count rows per finish
print("\nRows per finish:")
finish_counts = df['Option2 Value'].value_counts()
for finish, count in finish_counts.items():
    print(f"{finish}: {count} rows")

# Check SKU mapping
print("\nSKU to Finish mapping:")
# Group by Option2 Value (finish) and Variant SKU
sku_finish_map = {}
for idx, row in df.iterrows():
    finish = row['Option2 Value']
    sku = row['Variant SKU']
    if finish not in sku_finish_map:
        sku_finish_map[finish] = sku
    
for finish, sku in sku_finish_map.items():
    print(f"{finish} -> SKU: {sku}")

# Check groupings
print("\nFinishes by SKU grouping:")
skus = df['Variant SKU'].unique()
for sku in skus:
    finishes = df[df['Variant SKU'] == sku]['Option2 Value'].tolist()
    print(f"SKU {sku} ({len(finishes)} finishes):")
    for finish in finishes:
        print(f"  - {finish}")

# Check pricing
print("\nPricing by finish:")
for finish, sku in sku_finish_map.items():
    price = df[df['Option2 Value'] == finish]['Variant Price'].iloc[0]
    print(f"{finish} (SKU: {sku}): £{price}")

# Verify finishes are correctly mapped
print("\nVerifying finish mapping:")
scp_finishes = df[df['Option2 Value'] == 'Satin Chrome (SCP)']
pb_finishes = df[df['Option2 Value'] == 'Polished Brass (PB)']
cp_finishes = df[df['Option2 Value'] == 'Polished Chrome (CP)']

hash_codes = ["PN", "SN", "BZ", "AB", "SB", "DB", "BAB", "BZW", "BABW", "ABW", "DBW", "NBW", "SBW", "PBUL"]
xhash_codes = ["PCOP", "SCOP", "BLN", "PEW", "MBL", "ASV", "RGP", "ACOP"]

for code in hash_codes:
    # Check if any finishes contain this code
    matching_finishes = [f for f in df['Option2 Value'] if f"({code})" in f]
    if matching_finishes:
        # Get the SKU for this finish
        for finish in matching_finishes:
            sku = df[df['Option2 Value'] == finish]['Variant SKU'].iloc[0]
            price = df[df['Option2 Value'] == finish]['Variant Price'].iloc[0]
            print(f"## finish {finish} -> SKU: {sku}, Price: £{price}")

for code in xhash_codes:
    # Check if any finishes contain this code
    matching_finishes = [f for f in df['Option2 Value'] if f"({code})" in f]
    if matching_finishes:
        # Get the SKU for this finish
        for finish in matching_finishes:
            sku = df[df['Option2 Value'] == finish]['Variant SKU'].iloc[0]
            price = df[df['Option2 Value'] == finish]['Variant Price'].iloc[0]
            print(f"x## finish {finish} -> SKU: {sku}, Price: £{price}") 