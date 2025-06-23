import pandas as pd

# Load the Excel file
excel_file = 'MASTER COPY.xlsx'
print(f"Examining Excel file: {excel_file}")

# List all sheets in the Excel file
xls = pd.ExcelFile(excel_file)
sheets = xls.sheet_names
print("Sheets in the Excel file:")
for sheet in sheets:
    print(f"- {sheet}")

# Read the 'MASTER COPY' tab - just the first few rows to understand structure
print("\nSample data from 'MASTER COPY' tab:")
master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY', nrows=5)
print(master_copy_df.head())
print("\nColumns in 'MASTER COPY' tab:")
print(master_copy_df.columns.tolist())

# Read the sample rows mentioned in the requirements (14786 and 14787)
print("\nExamining specific rows 14786 and 14787 from 'MASTER COPY' tab:")
specific_rows = pd.read_excel(excel_file, sheet_name='MASTER COPY', skiprows=14785, nrows=2)
print(specific_rows)

# Check 'Sample' tab to see the expected output
print("\nSample data from 'Sample' tab:")
sample_df = pd.read_excel(excel_file, sheet_name='Sample', nrows=5)
print(sample_df.head())
print("\nColumns in 'Sample' tab:")
print(sample_df.columns.tolist())

# Check 'Mapping' tab to understand column mappings
print("\nChecking 'Mapping' tab:")
mapping_df = pd.read_excel(excel_file, sheet_name='Mapping', nrows=10)
print(mapping_df.head(10)) 