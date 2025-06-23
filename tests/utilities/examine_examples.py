import pandas as pd
import numpy as np

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

# Read the specific rows mentioned in the requirements (rows 14786 and 14787)
print("Examining specific rows 14786 and 14787 from 'MASTER COPY' tab:")
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
# Convert index to 1-based for matching row numbers
master_copy_df.index = master_copy_df.index + 1  
example_rows = master_copy_df.loc[[14786, 14787]]
print(example_rows)
print("\nColumns in example rows:")
print(example_rows.columns.tolist())

# Read the Sample tab to see the expected output
print("\nExamining 'Sample' tab (expected output):")
sample_df = pd.read_excel(excel_file, sheet_name='Sample')
print(f"Sample tab has {len(sample_df)} rows")
print(sample_df.head(5))

# Check if the Sample tab rows correspond to the example rows from MASTER COPY
print("\nCheck if example product code appears in Sample tab:")
example_product_code = example_rows['code'].iloc[0] if 'code' in example_rows.columns else None
if example_product_code:
    matching_rows = sample_df[sample_df['Variant SKU'] == example_product_code]
    print(f"Found {len(matching_rows)} rows in Sample tab matching product code {example_product_code}")
    if not matching_rows.empty:
        print(matching_rows.head(1))

# Let's try to understand the pattern using the product description
example_desc = example_rows['description'].iloc[0] if 'description' in example_rows.columns else None
if example_desc:
    print(f"\nSearching for products with description containing: {example_desc}")
    matching_desc = sample_df[sample_df['Title'].str.contains(str(example_desc), na=False)]
    print(f"Found {len(matching_desc)} rows in Sample tab with matching description")
    if not matching_desc.empty:
        print("\nFirst row of matching description in Sample tab:")
        print(matching_desc.iloc[0]) 