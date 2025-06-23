import pandas as pd
import numpy as np

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

# Read the specific rows mentioned in the requirements
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
example_rows = master_copy_df.loc[[14786, 14787]].copy()

print("Detailed view of example rows (14786-14787):")
print(example_rows)
print("\nColumn data types:")
for col in example_rows.columns:
    print(f"{col}: {example_rows[col].dtype}")

# Check for NaN values
print("\nNaN values in each column:")
for col in example_rows.columns:
    print(f"{col}: {example_rows[col].isna().tolist()}")

# Check if the second row has data in other columns
print("\nSecond row data:")
row2 = example_rows.iloc[1].dropna()
print(row2)

# Try to manually group the rows
print("\nManually grouping rows:")
# Fill forward the description to the second row if it's empty
if pd.isna(example_rows.iloc[1]['description']) and not pd.isna(example_rows.iloc[0]['description']):
    description = example_rows.iloc[0]['description']
    print(f"First row description: {description}")
    print(f"Setting this description for second row too")
    
    # Check for size values in both rows
    size1 = example_rows.iloc[0]['size']
    size2 = example_rows.iloc[1]['size']
    
    print(f"First row size: {size1}")
    print(f"Second row size: {size2}")
    
    # If both rows have different sizes, they should be treated as part of the same product
    if not pd.isna(size1) and not pd.isna(size2) and size1 != size2:
        print("Both rows have different sizes and should be treated as the same product")
else:
    print("Second row has its own description or both rows have no description")

# Load the sample tab to see the expected output
sample_df = pd.read_excel(excel_file, sheet_name='Sample')
print(f"\nSample tab has {len(sample_df)} rows")

# Check unique sizes in the sample
sizes = sample_df['Option1 Value'].unique()
print(f"Unique sizes in sample: {sizes}")

# Count rows per size
size_counts = sample_df.groupby('Option1 Value').size()
print("\nNumber of rows per size in sample:")
print(size_counts) 