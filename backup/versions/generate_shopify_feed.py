import pandas as pd
import numpy as np
import re
from datetime import datetime

def clean_string(s):
    """Clean a string to create a handle (lowercase, replace spaces with hyphens)"""
    if pd.isna(s):
        return ""
    return re.sub(r'[^a-zA-Z0-9\-]', '', str(s).lower().replace(' ', '-'))

def generate_shopify_feed(excel_file, output_file=None):
    """Generate a Shopify product feed from MASTER COPY tab for new products"""
    # Load the Excel file
    master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
    sample_df = pd.read_excel(excel_file, sheet_name='Sample')
    finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')
    
    # Create a template DataFrame for the Shopify feed using the columns from Sample tab
    template_columns = sample_df.columns.tolist()
    shopify_feed = pd.DataFrame(columns=template_columns)
    
    # Get test data (rows 14786 and 14787 as mentioned in the instructions)
    master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
    test_rows = master_copy_df.loc[[14786, 14787]].dropna(how='all')
    
    # Process each product (group of rows with the same description)
    product_groups = []
    current_group = []
    
    for _, row in test_rows.iterrows():
        if not pd.isna(row['description']):  # Start of a new product
            if current_group:
                product_groups.append(current_group)
                current_group = []
            current_group.append(row)
        elif current_group:  # Continuation of current product
            current_group.append(row)
    
    # Add the last group if it exists
    if current_group:
        product_groups.append(current_group)
    
    # Process each product group
    for product_group in product_groups:
        product_rows = []
        
        # Get product details from the first row
        first_row = product_group[0]
        product_description = first_row['description']
        
        # Generate handle from product description
        handle = clean_string(product_description)
        
        # Determine which finishes to use based on product name
        keywords = ['Bjorn', 'Cadiz', 'Denham', 'Wilton', 'Capri', 'Leon', 'Oxon']
        matching_keywords = [keyword for keyword in keywords if keyword.lower() in str(product_description).lower()]
        
        # Find finish column in Finishes tab
        finish_col = None
        if matching_keywords:
            keyword = matching_keywords[0]
            for col in finishes_df.columns:
                if str(col).lower().find(keyword.lower()) != -1:
                    finish_col = col
                    break
        
        # If no specific finish column found, check if finish count is specified
        if not finish_col and 'finish count' in first_row and not pd.isna(first_row['finish count']):
            finish_count = int(first_row['finish count'])
            matching_count_cols = [col for col in finishes_df.columns if str(col).startswith(str(finish_count))]
            if matching_count_cols:
                finish_col = matching_count_cols[0]
        
        # Default to the first column if no match found
        if not finish_col:
            # Look for columns that are purely numbers (finish counts)
            number_cols = [col for col in finishes_df.columns if str(col).isdigit()]
            if number_cols:
                finish_col = number_cols[0]
            else:
                finish_col = finishes_df.columns[0]
        
        # Get the finishes from the appropriate column
        finishes = finishes_df[finish_col].dropna().tolist()
        
        # Process each size variant
        for row in product_group:
            if pd.isna(row['size']):
                continue
                
            size = row['size']
            variant_sku = row['code']
            variant_price = row['rrp']
            
            # For each finish, create a row in the Shopify feed
            for i, finish in enumerate(finishes):
                # Create new row based on template
                new_row = {col: None for col in template_columns}
                
                # Set values based on mapping instructions
                new_row['Handle'] = handle
                
                # Only set certain fields for the first row of the product
                is_first_row = (i == 0 and row is product_group[0])
                
                if is_first_row:
                    new_row['Title'] = product_description
                    new_row['Vendor'] = "vendor-unknown"
                    new_row['Product Category'] = "Uncategorized"
                    new_row['Type'] = "Door Handles"  # This would need to be determined from the data
                    new_row['Published'] = 1.0
                    new_row['Option1 Name'] = "Size"
                    new_row['Option2 Name'] = "Finish"
                    new_row['Image Src'] = sample_df['Image Src'].iloc[0]  # Use an example from sample
                    new_row['Image Position'] = 1
                    new_row['Gift Card'] = "FALSE"
                    new_row['SEO Title'] = f"{product_description} | A&H Brass"
                    
                    # Set regional inclusion
                    for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                        new_row[f'Included / {region}'] = 1.0
                    
                    new_row['Status'] = "active"
                
                # Set variant-specific values
                new_row['Option1 Value'] = size
                new_row['Option2 Value'] = finish
                new_row['Variant SKU'] = variant_sku
                new_row['Variant Grams'] = 0
                new_row['Variant Inventory Tracker'] = "shopify"
                new_row['Variant Inventory Qty'] = 10000
                new_row['Variant Inventory Policy'] = "continue"
                new_row['Variant Fulfillment Service'] = "manual"
                new_row['Variant Price'] = variant_price
                new_row['Variant Requires Shipping'] = "TRUE"
                new_row['Variant Taxable'] = "TRUE"
                new_row['Variant Image'] = sample_df['Image Src'].iloc[0] if pd.notna(sample_df['Image Src'].iloc[0]) else None
                new_row['Variant Weight Unit'] = "kg"
                
                # Add the row to our product rows
                product_rows.append(new_row)
        
        # Add all product rows to the Shopify feed
        for row in product_rows:
            shopify_feed = pd.concat([shopify_feed, pd.DataFrame([row])], ignore_index=True)
    
    # If output file is specified, save the feed
    if output_file:
        shopify_feed.to_excel(output_file, index=False)
        print(f"Shopify feed saved to {output_file}")
    
    return shopify_feed

if __name__ == "__main__":
    # Generate test feed using example rows
    excel_file = 'MASTER COPY.xlsx'
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f'shopify_feed_test_{timestamp}.xlsx'
    
    feed = generate_shopify_feed(excel_file, output_file)
    print(f"Generated {len(feed)} rows in the Shopify feed")
    
    # Print sample of the feed
    print("\nSample of generated Shopify feed:")
    sample_columns = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
    print(feed[sample_columns].head(10)) 