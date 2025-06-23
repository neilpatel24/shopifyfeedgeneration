import pandas as pd
import numpy as np

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

# Examine the Finishes tab
finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')
print("Finishes tab structure:")
print(finishes_df.head(10))
print(f"\nColumns in Finishes tab: {finishes_df.columns.tolist()}")

# Check the example rows from MASTER COPY
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
example_rows = master_copy_df.loc[[14786, 14787]]

# Load the Sample tab to see which finishes were used
sample_df = pd.read_excel(excel_file, sheet_name='Sample')

# Group by size and count finishes
sample_sizes = sample_df['Option1 Value'].unique()
print(f"\nUnique sizes in the Sample tab: {sample_sizes}")

# Count rows per size
size_counts = sample_df.groupby('Option1 Value').size()
print("\nNumber of rows per size:")
print(size_counts)

# Check unique finishes
unique_finishes = sample_df['Option2 Value'].unique()
print(f"\nUnique finishes in the Sample tab ({len(unique_finishes)}):")
for i, finish in enumerate(unique_finishes):
    print(f"{i+1}. {finish}")

# Check if the product name in the example contains keywords to determine which finish column to use
example_description = example_rows['description'].iloc[0] if not pd.isna(example_rows['description'].iloc[0]) else ""
keywords = ['Bjorn', 'Cadiz', 'Denham', 'Wilton', 'Capri', 'Leon', 'Oxon']
matching_keywords = [keyword for keyword in keywords if keyword.lower() in str(example_description).lower()]

print(f"\nProduct description: {example_description}")
print(f"Matching keywords for finish column selection: {matching_keywords}")

# Determine which column in Finishes tab to use based on keywords
if matching_keywords:
    keyword = matching_keywords[0]
    print(f"\nProduct contains keyword '{keyword}', checking for corresponding column in Finishes tab")
    
    # Case-insensitive matching for column names
    matching_cols = []
    for col in finishes_df.columns:
        if str(col).lower().find(keyword.lower()) != -1:
            matching_cols.append(col)
    
    print(f"Matching columns in Finishes tab: {matching_cols}")
    
    # If a matching column is found, show its finishes
    if matching_cols:
        finish_col = matching_cols[0]
        print(f"\nFinishes from column '{finish_col}':")
        finishes = finishes_df[finish_col].dropna().tolist()
        for i, finish in enumerate(finishes):
            print(f"{i+1}. {finish}")
            
        # Compare with actual finishes used in Sample tab
        print("\nComparing with finishes used in Sample tab:")
        for finish in finishes:
            found = any(finish in str(sample_finish) for sample_finish in unique_finishes)
            print(f"{finish}: {'Found' if found else 'Not found'}")
    else:
        print("\nNo matching columns found in Finishes tab for keyword:", keyword)
else:
    print("\nNo matching keywords found in product description for finish column selection")

# Get finish count from example rows
print("\nChecking finish count in example rows:")
if 'finish count' in example_rows.columns:
    finish_count = example_rows['finish count'].iloc[0]
    print(f"Finish count in example row: {finish_count}")
    if not pd.isna(finish_count):
        # Look for column with matching count
        matching_count_cols = [col for col in finishes_df.columns if str(col).startswith(str(int(finish_count)))]
        print(f"Columns matching finish count {int(finish_count)}: {matching_count_cols}")
        
        if matching_count_cols:
            count_col = matching_count_cols[0]
            print(f"\nFinishes from column with count {count_col}:")
            count_finishes = finishes_df[count_col].dropna().tolist()
            for i, finish in enumerate(count_finishes):
                print(f"{i+1}. {finish}")
else:
    print("No 'finish count' column found in example rows")
    
# Check for specific finishes mentioned in the example rows
if 'finish' in example_rows.columns:
    example_finish = example_rows['finish'].iloc[0]
    print(f"\nFinish specified in example row: {example_finish}")
    
    # If finish contains '##', it indicates multiple finishes
    if example_finish and '##' in str(example_finish):
        print("'##' found in finish, indicating multiple finishes available") 