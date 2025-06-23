import pandas as pd
import openpyxl

# First check the raw data to verify there are multiple products in rows 14786-14812
print("Examining raw data for rows 14786-14812:")
try:
    # Check with openpyxl first (most accurate row detection)
    wb = openpyxl.load_workbook('MASTER COPY.xlsx', read_only=True)
    sheet = wb['MASTER COPY']
    
    # Gather all descriptions to identify product boundaries
    descriptions = {}
    for row_num in range(14786, 14813):
        # Find the description column (typically column 4)
        desc = sheet.cell(row=row_num, column=4).value
        if desc is not None and isinstance(desc, str) and len(desc) > 5:
            descriptions[row_num] = desc
    
    wb.close()
    
    # Print all found descriptions to identify different products
    print("\nFound product descriptions:")
    current_desc = None
    product_groups = []
    current_group = []
    
    for row, desc in sorted(descriptions.items()):
        print(f"Row {row}: {desc}")
        
        # Check if this is a new product description
        if current_desc != desc:
            if current_group:
                product_groups.append((current_desc, current_group))
            current_desc = desc
            current_group = [row]
        else:
            current_group.append(row)
    
    # Add the last group
    if current_group:
        product_groups.append((current_desc, current_group))
    
    # Report how many different products were found
    print(f"\nIdentified {len(product_groups)} distinct products:")
    for i, (desc, rows) in enumerate(product_groups):
        print(f"Product {i+1}: {desc}")
        print(f"  Rows: {rows}")
        print(f"  Row count: {len(rows)}")
    
    # Now run a test to see if the script correctly separates products
    from shopify_feed_generator import CONFIG, generate_shopify_feed
    
    print("\nTesting script with rows 14786-14812 to check product separation:")
    
    # Set the test range
    CONFIG["test_start_row"] = 14786
    CONFIG["test_end_row"] = 14812
    
    # Check how products would be grouped in the current code
    master_copy_df = pd.read_excel('MASTER COPY.xlsx', sheet_name='MASTER COPY')
    master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
    
    # Get rows in our test range
    test_rows = master_copy_df.loc[14786:14812].copy()
    
    # How current grouping would work
    def group_products(df):
        """Group products by description to handle multiple sizes of the same product"""
        product_groups = []
        current_group = []
        current_description = None
        
        for idx, row in df.iterrows():
            # If the row has a description and it's different from the current one, start a new group
            if not pd.isna(row['description']):
                if current_description != row['description']:
                    if current_group:
                        product_groups.append(current_group)
                    current_group = [row]
                    current_description = row['description']
                else:
                    # Same description, add to current group
                    current_group.append(row)
            elif current_group:
                # No description, but might be a continuation of current product
                current_group.append(row)
        
        # Add the last group if it exists
        if current_group:
            product_groups.append(current_group)
        
        return product_groups
    
    # Test our grouping function
    groups = group_products(test_rows)
    print(f"\nCurrent grouping function identified {len(groups)} groups:")
    
    for i, group in enumerate(groups):
        # Just print the first row's description and index
        if len(group) > 0:
            first_row = group[0]
            description = first_row['description'] if 'description' in first_row and not pd.isna(first_row['description']) else "No description"
            print(f"Group {i+1}: {description}")
            print(f"  Row count: {len(group)}")
            
            # Print the indices of the first few rows
            indices = []
            for j in range(min(5, len(group))):
                indices.append(group[j].name)  # .name gives the index of the row
            print(f"  First few rows: {indices}")
    
    # Now run the actual script to see how it handles the test range
    print("\nRunning generate_shopify_feed on the test range:")
    feed = generate_shopify_feed('MASTER COPY.xlsx', 'test_product_separation.xlsx', test_mode=True)
    
    # Analyze the output to see how products were grouped
    print(f"\nGenerated {len(feed)} rows in total")
    
    # Count how many unique products are in the output
    unique_handles = feed['Handle'].unique()
    print(f"Found {len(unique_handles)} unique product handles:")
    for handle in unique_handles:
        product_rows = feed[feed['Handle'] == handle]
        # Get the title from the first row
        title = product_rows['Title'].dropna().iloc[0] if not product_rows['Title'].dropna().empty else "Unknown"
        print(f"  {handle}: {title} ({len(product_rows)} variants)")
    
except Exception as e:
    print(f"Error: {e}") 