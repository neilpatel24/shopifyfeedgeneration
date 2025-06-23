import pandas as pd

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'

try:
    # Read the MASTER COPY tab
    df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
    
    # Print basic info
    print(f"Total rows in MASTER COPY: {len(df)}")
    print(f"Index starts at: {df.index[0]}")
    print(f"Index ends at: {df.index[-1]}")
    
    # Convert to 1-based indexing
    df.index = df.index + 1
    print(f"After 1-based conversion, index starts at: {df.index[0]}")
    print(f"After 1-based conversion, index ends at: {df.index[-1]}")
    
    # Check if specific rows exist
    check_rows = [14769, 14770, 14790, 14812]
    print("\nChecking if specific rows exist:")
    
    for row in check_rows:
        if row in df.index:
            print(f"Row {row} exists")
            # Show some data from this row
            row_data = df.loc[row]
            description = row_data.get('description', 'N/A')
            print(f"  Description: {description}")
        else:
            print(f"Row {row} does NOT exist")
    
    # Check for the first description after row 14790
    print("\nChecking for rows after 14700:")
    for i in range(14700, 14850):
        if i in df.index and not pd.isna(df.loc[i].get('description')):
            print(f"Row {i} has description: {df.loc[i]['description']}")

except Exception as e:
    print(f"Error reading the Excel file: {e}") 