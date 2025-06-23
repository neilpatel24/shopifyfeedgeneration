import pandas as pd

# Load the output file
df = pd.read_excel('test_correct_status.xlsx')

# Check the Status column
if 'Status' in df.columns:
    status_values = df['Status'].unique()
    print(f"Status column values: {status_values}")
    
    # Count occurrences
    status_counts = df['Status'].value_counts(dropna=False)
    print("\nStatus value counts:")
    for status, count in status_counts.items():
        status_str = str(status) if not pd.isna(status) else "NaN"
        print(f"{status_str}: {count} rows")
    
    # Check if status is only on the first row
    print("\nChecking if Status is only on first row:")
    # Get first row with Status
    first_status_row = df[~df['Status'].isna()].iloc[0] if any(~df['Status'].isna()) else None
    if first_status_row is not None:
        print(f"First row with Status value: {first_status_row[['Handle', 'Title', 'Option1 Value', 'Status']]}")
    
    # Print sample rows with Status
    print("\nSample rows:")
    sample_columns = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Status']
    print(df[sample_columns].head(10))
else:
    print("Status column not found in the file") 