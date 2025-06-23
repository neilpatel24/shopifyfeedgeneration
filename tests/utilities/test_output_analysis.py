import pandas as pd

# Load the generated Excel file
df = pd.read_excel('test_multiple_products.xlsx')
print(f'Total rows: {len(df)}')

# Check how many unique products are in the file
unique_handles = df['Handle'].unique()
print(f'Unique products: {len(unique_handles)}')

# Print details for each product
for handle in unique_handles:
    product_rows = df[df['Handle'] == handle]
    
    # Get the title from the first row that has one
    title = "Unknown"
    title_rows = product_rows['Title'].dropna()
    if not title_rows.empty:
        title = title_rows.iloc[0]
    
    # Count unique sizes and finishes
    unique_sizes = product_rows['Option1 Value'].unique()
    unique_finishes = product_rows['Option2 Value'].unique()
    
    print(f'  {handle}: {title}')
    print(f'    Variants: {len(product_rows)}')
    print(f'    Unique sizes: {len(unique_sizes)}')
    print(f'    Unique finishes: {len(unique_finishes)}') 