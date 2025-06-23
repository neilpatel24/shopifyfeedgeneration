import pandas as pd

# Load the output file
df = pd.read_excel('test_inventory_policy.xlsx')

# Check the Variant Inventory Policy column
if 'Variant Inventory Policy' in df.columns:
    policy_values = df['Variant Inventory Policy'].unique()
    print(f"Variant Inventory Policy values: {policy_values}")
    
    # Count occurrences
    policy_counts = df['Variant Inventory Policy'].value_counts()
    print("\nVariant Inventory Policy value counts:")
    for policy, count in policy_counts.items():
        print(f"{policy}: {count} rows")
else:
    print("Variant Inventory Policy column not found in the file")

# Check a sample of rows
print("\nSample rows with Variant Inventory Policy:")
if 'Variant Inventory Policy' in df.columns:
    sample_columns = ['Handle', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Inventory Policy']
    print(df[sample_columns].head(5))
else:
    print("Variant Inventory Policy column not found in the file") 