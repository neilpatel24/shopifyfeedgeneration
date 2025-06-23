import pandas as pd
import numpy as np
import re

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

# Load the Sample tab to understand how SKUs are structured
sample_df = pd.read_excel(excel_file, sheet_name='Sample')
print("Analyzing SKU structure in Sample tab:")

# Check the SKUs for each variant
variant_skus = sample_df[['Option1 Value', 'Option2 Value', 'Variant SKU']]
print(variant_skus.head(16))

# Examine the supplier code pattern in MASTER COPY
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
example_rows = master_copy_df.loc[[14786, 14787]]

print("\nExample rows from MASTER COPY:")
print(example_rows[['code', 'supp', 'supp code', 'description', 'size', 'finish']])

# Check if supplier code has a pattern with ## for finishes
if 'supp code' in example_rows.columns:
    supp_code = example_rows['supp code'].iloc[0]
    print(f"\nSupplier code in example: {supp_code}")
    
    # Check if the code has a pattern with ## for finishes
    if supp_code and '##' in str(supp_code):
        print("'##' found in supplier code, indicating a placeholder for finish codes")
        
        # Check the pattern of the supplier code
        pattern = re.sub(r'##', r'(.*)', str(supp_code))
        print(f"Supplier code pattern: {pattern}")

# Compare this with the actual SKUs in the Sample tab
print("\nComparing with SKUs in Sample tab:")
if 'code' in example_rows.columns:
    example_code = example_rows['code'].iloc[0]
    matching_skus = sample_df[sample_df['Variant SKU'] == example_code]
    print(f"Found {len(matching_skus)} SKUs matching code {example_code}")
    
    if not matching_skus.empty:
        print("\nExample SKU entries:")
        print(matching_skus[['Option1 Value', 'Option2 Value', 'Variant SKU']].head(5))

# Check relationship between MASTER COPY and Sample tab
print("\nAnalyzing relationship between MASTER COPY and Sample tab:")
if 'size' in example_rows.columns:
    example_size = example_rows['size'].iloc[0]
    print(f"Size in MASTER COPY: {example_size}")
    
    # Check if this size appears in Sample tab
    matching_size = sample_df[sample_df['Option1 Value'] == example_size]
    print(f"Found {len(matching_size)} entries with matching size in Sample tab")

# Check if all finishes in the corresponding column of Finishes tab are used
print("\nChecking if all finishes from Finishes tab are used in Sample tab:")
finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')

# Find the cadiz column
cadiz_col = next((col for col in finishes_df.columns if 'cadiz' in str(col).lower()), None)
if cadiz_col:
    cadiz_finishes = finishes_df[cadiz_col].dropna().tolist()
    print(f"Found {len(cadiz_finishes)} finishes in '{cadiz_col}' column:")
    
    # Check if all these finishes appear in the Sample tab
    sample_finishes = sample_df['Option2 Value'].unique()
    for finish in cadiz_finishes:
        found = any(finish in str(sample_finish) for sample_finish in sample_finishes)
        print(f"  {finish}: {'Found' if found else 'Not found'} in Sample tab")
    
    # Check if Sample tab has the exact same number of finishes per size
    sizes = sample_df['Option1 Value'].unique()
    for size in sizes:
        size_finishes = sample_df[sample_df['Option1 Value'] == size]['Option2 Value'].tolist()
        print(f"\nSize {size} has {len(size_finishes)} finishes:")
        for finish in size_finishes:
            print(f"  {finish}") 