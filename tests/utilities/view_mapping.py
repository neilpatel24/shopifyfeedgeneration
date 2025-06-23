import pandas as pd
import numpy as np

# Load the mapping tab to understand the column requirements
excel_file = 'MASTER COPY.xlsx'
mapping_df = pd.read_excel(excel_file, sheet_name='Mapping')

# Display the mapping instructions
pd.set_option('display.max_colwidth', None)  # Show full column content
print("Mapping Instructions:")
for _, row in mapping_df.iterrows():
    if not pd.isna(row.iloc[0]) and not pd.isna(row.iloc[1]):
        print(f"{row.iloc[0]}: {row.iloc[1]}")
        print("-" * 100)

# Also load the Finishes tab to understand how finishes work
finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')
print("\nFinishes Information:")
print(finishes_df.head(10))

# Now examine the sample data more carefully
print("\nSample output data:")
sample_df = pd.read_excel(excel_file, sheet_name='Sample')
print(f"Sample has {len(sample_df)} rows")
print(sample_df.head(3)) 