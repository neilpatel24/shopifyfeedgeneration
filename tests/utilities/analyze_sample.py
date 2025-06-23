import pandas as pd
import numpy as np

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

# Read the specific rows mentioned in the requirements
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
example_rows = master_copy_df.loc[[14786, 14787]]
print(f"Example rows from MASTER COPY (rows 14786-14787):")
print(example_rows)

# Read the Sample tab that shows the expected output
sample_df = pd.read_excel(excel_file, sheet_name='Sample')

# Find the product in the Sample tab
example_product_code = example_rows['code'].iloc[0]
matching_rows = sample_df[sample_df['Variant SKU'] == example_product_code]
print(f"\nFound {len(matching_rows)} rows in Sample tab for product code {example_product_code}")

# Get product description from example rows
example_description = example_rows['description'].iloc[0] if not pd.isna(example_rows['description'].iloc[0]) else None
print(f"\nProduct description from MASTER COPY: {example_description}")

# Show the sizes and finishes in the example rows
example_sizes = example_rows['size'].tolist()
example_finishes_count = example_rows['finish count'].iloc[0] if 'finish count' in example_rows.columns and not pd.isna(example_rows['finish count'].iloc[0]) else None
print(f"Sizes in example rows: {example_sizes}")
print(f"Finish count in example rows: {example_finishes_count}")

# Now analyze the Sample tab structure in detail
print("\nDetailed analysis of the Sample tab structure:")
print(f"Total rows in Sample tab: {len(sample_df)}")

# Check how many different products are in the Sample tab
unique_handles = sample_df['Handle'].nunique()
print(f"Number of unique product handles: {unique_handles}")

# Look at the first unique product
first_product_handle = sample_df['Handle'].iloc[0]
first_product_rows = sample_df[sample_df['Handle'] == first_product_handle]
print(f"\nAnalyzing first product with handle: {first_product_handle}")
print(f"Number of rows for this product: {len(first_product_rows)}")

# Show the sizes and finishes for this product
sizes = first_product_rows['Option1 Value'].tolist()
finishes = first_product_rows['Option2 Value'].tolist()
print(f"Sizes for this product: {sizes}")
print(f"Finishes for this product: {finishes}")

# Check how variants are structured
print("\nSample of variant structure:")
variant_sample = first_product_rows[['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price', 'Published']].head(8)
print(variant_sample)

# Look for differences between first row and subsequent rows for the same product
first_row = first_product_rows.iloc[0].to_dict()
diff_columns = []

for col in sample_df.columns:
    is_different = False
    for i in range(1, len(first_product_rows)):
        if first_product_rows.iloc[i][col] != first_row[col] and not (pd.isna(first_product_rows.iloc[i][col]) and pd.isna(first_row[col])):
            is_different = True
            break
    if is_different:
        diff_columns.append(col)

print("\nColumns that differ between first row and subsequent rows:")
print(diff_columns) 