import pandas as pd
import numpy as np
import re
import argparse
from datetime import datetime
import warnings

# Suppress the FutureWarning about DataFrame concatenation
warnings.simplefilter(action='ignore', category=FutureWarning)

# Configuration for test rows
CONFIG = {
    "test_start_row": 14786,  # Default start row
    "test_end_row": 14787     # Default end row
}

def clean_string(s):
    """Clean a string to create a handle (lowercase, replace spaces with hyphens)"""
    if pd.isna(s):
        return ""
    return re.sub(r'[^a-zA-Z0-9\-]', '', str(s).lower().replace(' ', '-'))

def find_new_products(master_copy_df, existing_feed_df=None):
    """Find new products in the MASTER COPY tab that need to be added to the feed"""
    # If no existing feed is provided, all products are considered new
    if existing_feed_df is None or existing_feed_df.empty:
        return master_copy_df
    
    # Get all SKUs in the existing feed
    existing_skus = set(existing_feed_df['Variant SKU'].dropna().unique())
    
    # Find rows in master_copy that have SKUs not in the existing feed
    new_product_rows = master_copy_df[~master_copy_df['code'].isin(existing_skus)]
    
    return new_product_rows

def group_products(master_copy_df):
    """Group products by description to handle multiple sizes of the same product"""
    product_groups = []
    current_group = []
    current_description = None
    
    for _, row in master_copy_df.iterrows():
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

def get_product_type(description):
    """Determine the product type based on description keywords"""
    description = str(description).lower()
    
    if any(term in description for term in ['handle', 'lever', 'knob']):
        return "Door Handles"
    elif any(term in description for term in ['bathroom', 'shower', 'toilet', 'bath']):
        return "Bathroom"
    elif any(term in description for term in ['tube', 'fitting']):
        return "Tube Fittings"
    else:
        return "Miscellaneous"

def get_finishes_for_product(product_description, finish_count, finishes_df):
    """Determine which finishes to use for a product"""
    # Check if product name contains keywords to determine which finish column to use
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
    if not finish_col and not pd.isna(finish_count):
        try:
            finish_count = int(finish_count)
            matching_count_cols = [col for col in finishes_df.columns if str(col).startswith(str(finish_count))]
            if matching_count_cols:
                finish_col = matching_count_cols[0]
        except (ValueError, TypeError):
            pass
    
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
    return finishes

def generate_shopify_feed(excel_file, output_file=None, test_mode=False):
    """Generate a Shopify product feed from MASTER COPY tab for new products"""
    # Load the Excel file
    master_copy_df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
    sample_df = pd.read_excel(excel_file, sheet_name='Sample')
    finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')
    
    # Create a template DataFrame for the Shopify feed using the columns from Sample tab
    template_columns = sample_df.columns.tolist()
    shopify_feed = pd.DataFrame(columns=template_columns)
    
    # If test_mode is True, only use the specified rows
    if test_mode:
        # Use either the default rows (14786-14787) or custom rows if provided
        start_row = CONFIG["test_start_row"]
        end_row = CONFIG["test_end_row"]
        
        print(f"Running in test mode with rows {start_row}-{end_row}")
        
        # Get all rows from MASTER COPY for product details
        master_copy_df.index = master_copy_df.index + 1  # Convert to 1-based indexing
        
        # Include the row above as a reference if it's not the first row
        if start_row > 1:
            header_row = start_row - 1
            product_rows = master_copy_df.loc[header_row:end_row].copy()
        else:
            product_rows = master_copy_df.loc[start_row:end_row].copy()
        
        # Filter out rows with no description or empty rows
        valid_rows = product_rows[~pd.isna(product_rows['description'])].copy()
        
        if valid_rows.empty:
            print("Error: No valid product rows found in the specified range")
            return pd.DataFrame()
        
        # Get product description from the first valid row
        product_description = valid_rows.iloc[0]['description']
        print(f"Product: {product_description}")
        
        # Get all valid sizes, SKUs, and prices from the rows
        size_data = {}
        for idx, row in valid_rows.iterrows():
            if not pd.isna(row['size']) and not pd.isna(row['code']) and not pd.isna(row['rrp']):
                size = row['size']
                sku = str(row['code'])
                price = float(row['rrp'])
                
                # Store data by row index to preserve each unique SKU/price combination
                size_data[idx] = {
                    "size": size,
                    "price": price,
                    "sku": sku
                }
                print(f"Row {idx}: Size={size}, SKU={sku}, Price=£{price}")
        
        if not size_data:
            print("Error: No valid size/SKU/price data found in rows")
            return pd.DataFrame()
        
        print(f"Size data: {size_data}")
        
        # Determine the appropriate finishes column
        # Look for '##' in any of the valid rows' finish column
        finish_column = None
        for idx, row in valid_rows.iterrows():
            if not pd.isna(row['finish']) and '##' in str(row['finish']):
                # Check if finish count is specified
                if not pd.isna(row['finish count']):
                    finish_count = int(row['finish count'])
                    print(f"Found finish count: {finish_count}")
                    # Look for a column with this number
                    for col in finishes_df.columns:
                        if str(finish_count) == str(col) or str(col).startswith(str(finish_count) + ' '):
                            finish_column = col
                            break
                else:
                    # If no finish count, look for columns with specific names
                    # Check product name for keywords like Bjorn, Cadiz, etc.
                    keywords = ['Bjorn', 'Cadiz', 'Denham', 'Wilton', 'Capri', 'Leon', 'Oxon']
                    for keyword in keywords:
                        if keyword.lower() in product_description.lower():
                            for col in finishes_df.columns:
                                if keyword.lower() in str(col).lower():
                                    finish_column = col
                                    break
                            if finish_column:
                                break
        
        # If still no finish column found, check if the row has 'x##'
        if not finish_column:
            for idx, row in valid_rows.iterrows():
                if not pd.isna(row['finish']) and 'x##' in str(row['finish']):
                    # This indicates all finishes should be used - use column 25
                    for col in finishes_df.columns:
                        if str(col) == '25' or str(col) == 25:
                            finish_column = col
                            break
        
        # If still no column found, use the first numeric column
        if not finish_column:
            for col in finishes_df.columns:
                if str(col).isdigit() or str(col).split(' ')[0].isdigit():
                    finish_column = col
                    break
        
        # If still no column found, use the first column
        if not finish_column:
            finish_column = finishes_df.columns[0]
        
        print(f"Using finishes from column: {finish_column}")
        finishes = finishes_df[finish_column].dropna().tolist()
        print(f"Using {len(finishes)} finishes: {finishes}")
        
        # Create rows for each size-finish combination
        product_rows = []
        
        # Track if we've already set the first-row-only fields
        first_row_set = False
        
        # Process each product row (with unique SKU/price)
        for row_idx, data in size_data.items():
            size = data["size"]
            sku_base = data["sku"]
            price = data["price"]
            
            # For each finish, create a row in the Shopify feed
            for i, finish in enumerate(finishes):
                # Create new row based on template
                new_row = {col: None for col in template_columns}
                
                # Set values based on mapping instructions
                new_row['Handle'] = clean_string(product_description)
                
                # Only set certain fields for the first row of the product
                is_first_row = not first_row_set
                
                if is_first_row:
                    new_row['Title'] = product_description
                    new_row['Vendor'] = "vendor-unknown"
                    new_row['Product Category'] = "Uncategorized"
                    new_row['Type'] = get_product_type(product_description)
                    # Use string "TRUE" instead of boolean True
                    new_row['Published'] = "TRUE"
                    new_row['Option1 Name'] = "Size"
                    new_row['Option2 Name'] = "Finish"
                    
                    # Use an example image from sample
                    if not sample_df.empty and 'Image Src' in sample_df.columns and not pd.isna(sample_df['Image Src'].iloc[0]):
                        new_row['Image Src'] = sample_df['Image Src'].iloc[0]
                        
                    new_row['Image Position'] = 1
                    new_row['Gift Card'] = "FALSE"
                    new_row['SEO Title'] = f"{product_description} | A&H Brass"
                    
                    # Set regional inclusion - use "TRUE" string instead of boolean True
                    for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                        new_row[f'Included / {region}'] = "TRUE"
                    
                    new_row['Status'] = "draft"
                    
                    # Mark that we've set the first row fields
                    first_row_set = True
                
                # Set variant-specific values
                new_row['Option1 Value'] = size
                new_row['Option2 Value'] = finish
                
                # Set SKU - if it's the special 'X' variant, append the finish code
                variant_sku = sku_base
                
                # Variant price is the exact price from the Excel file
                new_row['Variant SKU'] = variant_sku
                new_row['Variant Grams'] = 0
                new_row['Variant Inventory Tracker'] = "shopify"
                new_row['Variant Inventory Qty'] = 10000
                new_row['Variant Inventory Policy'] = "continue"
                new_row['Variant Fulfillment Service'] = "manual"
                new_row['Variant Price'] = price
                
                # Set these as string literals
                new_row['Variant Requires Shipping'] = "TRUE"
                new_row['Variant Taxable'] = "TRUE"
                
                # Use the same image for all variants
                if not sample_df.empty and 'Image Src' in sample_df.columns and not pd.isna(sample_df['Image Src'].iloc[0]):
                    new_row['Variant Image'] = sample_df['Image Src'].iloc[0]
                    
                new_row['Variant Weight Unit'] = "kg"
                
                # Add the row to our product rows
                product_rows.append(new_row)
        
        # Add all product rows to the Shopify feed
        for row in product_rows:
            new_df = pd.DataFrame([row])
            shopify_feed = pd.concat([shopify_feed, new_df], ignore_index=True)
    
    else:
        # Normal processing for non-test mode
        # Find new products
        try:
            existing_feed_df = pd.read_excel(excel_file, sheet_name='ExampleFeed')
            new_products_df = find_new_products(master_copy_df, existing_feed_df)
            print(f"Found {len(new_products_df)} new products to add")
        except Exception as e:
            print(f"Could not load existing feed: {e}")
            new_products_df = master_copy_df
        
        # Group products by description
        product_groups = group_products(new_products_df)
        print(f"Grouped into {len(product_groups)} product sets")
        
        # Process each product group
        for product_group in product_groups:
            product_rows = []
            
            # Get product details from the first row
            first_row = product_group[0]
            product_description = first_row['description']
            
            # Skip if no description
            if pd.isna(product_description):
                continue
            
            print(f"Processing product: {product_description}")
            
            # Generate handle from product description
            handle = clean_string(product_description)
            
            # Determine product type
            product_type = get_product_type(product_description)
            
            # Get finishes for this product
            finish_count = first_row.get('finish count', None)
            finishes = get_finishes_for_product(product_description, finish_count, finishes_df)
            
            # Process each size variant
            for row in product_group:
                if pd.isna(row.get('size')):
                    continue
                    
                size = row['size']
                variant_sku = row['code']
                base_price = float(row['rrp']) if not pd.isna(row['rrp']) else 0.0
                
                print(f"  Size: {size}, SKU: {variant_sku}, Base Price: {base_price}")
                
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
                        new_row['Type'] = product_type
                        # Use string "TRUE" instead of boolean True
                        new_row['Published'] = "TRUE"
                        
                        # Check if this is a lever handle on plate for Option1 Name
                        if any(tag in product_description.lower() for tag in ["lever handles on plate", "lever handle on plate"]):
                            new_row['Option1 Name'] = "Option"
                        else:
                            new_row['Option1 Name'] = "Size"
                            
                        new_row['Option2 Name'] = "Finish"
                        
                        # Use an example image from sample
                        if not sample_df.empty and 'Image Src' in sample_df.columns and not pd.isna(sample_df['Image Src'].iloc[0]):
                            new_row['Image Src'] = sample_df['Image Src'].iloc[0]
                            
                        new_row['Image Position'] = 1
                        new_row['Gift Card'] = "FALSE"
                        new_row['SEO Title'] = f"{product_description} | A&H Brass"
                        
                        # Set regional inclusion - use "TRUE" string instead of boolean True
                        for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                            new_row[f'Included / {region}'] = "TRUE"
                        
                        new_row['Status'] = "draft"
                    
                    # Set variant-specific values
                    new_row['Option1 Value'] = size
                    new_row['Option2 Value'] = finish
                    new_row['Variant SKU'] = variant_sku
                    new_row['Variant Grams'] = 0
                    new_row['Variant Inventory Tracker'] = "shopify"
                    new_row['Variant Inventory Qty'] = 10000
                    new_row['Variant Inventory Policy'] = "continue"
                    new_row['Variant Fulfillment Service'] = "manual"
                    
                    # Identify premium finishes - these will cost more
                    premium_keywords = ['Polished Nickel', 'Brushed Nickel', 'Antique', 'FFPN', 'FFSN', 'FFAB']
                    
                    # Explicitly check each finish against premium keywords
                    is_premium = False
                    for keyword in premium_keywords:
                        if keyword in finish:
                            is_premium = True
                            break
                    
                    if is_premium:
                        # Premium finishes cost 10% more
                        variant_price = base_price * 1.1
                    else:
                        variant_price = base_price
                        
                    new_row['Variant Price'] = variant_price
                    # Set these as string literals
                    new_row['Variant Requires Shipping'] = "TRUE"
                    new_row['Variant Taxable'] = "TRUE"
                    
                    # Use the same image for all variants
                    if not sample_df.empty and 'Image Src' in sample_df.columns and not pd.isna(sample_df['Image Src'].iloc[0]):
                        new_row['Variant Image'] = sample_df['Image Src'].iloc[0]
                        
                    new_row['Variant Weight Unit'] = "kg"
                    
                    # Add the row to our product rows
                    product_rows.append(new_row)
            
            # Add all product rows to the Shopify feed
            for row in product_rows:
                new_df = pd.DataFrame([row])
                shopify_feed = pd.concat([shopify_feed, new_df], ignore_index=True)
    
    # Explicitly convert boolean columns to string literals "TRUE" or "FALSE"
    boolean_columns = ['Published', 'Variant Requires Shipping', 'Variant Taxable', 'Gift Card']
    for col in boolean_columns:
        if col in shopify_feed.columns:
            shopify_feed[col] = shopify_feed[col].apply(
                lambda x: "TRUE" if (x is True or x == 1 or x == "TRUE" or x == "True" or x == 1.0) 
                else ("FALSE" if (x is False or x == 0 or x == "FALSE" or x == "False" or x == 0.0) 
                else x)
            )
    
    # Make sure all regional inclusion columns use "TRUE"
    inclusion_columns = [col for col in shopify_feed.columns if col.startswith('Included /')]
    for col in inclusion_columns:
        shopify_feed[col] = shopify_feed[col].apply(
            lambda x: "TRUE" if (x is True or x == 1 or x == "TRUE" or x == "True" or x == 1.0) 
            else x
        )
    
    # If output file is specified, save the feed
    if output_file:
        # Save to temporary file to prevent pandas from converting "TRUE"/"FALSE" to boolean
        temp_file = f"temp_{output_file}"
        shopify_feed.to_excel(temp_file, index=False)
        
        # Read back the file and ensure boolean columns are strings
        df = pd.read_excel(temp_file)
        for col in boolean_columns + inclusion_columns:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda x: "TRUE" if (x is True or x == 1 or x == "TRUE" or x == "True" or x == 1.0) 
                    else ("FALSE" if (x is False or x == 0 or x == "FALSE" or x == "False" or x == 0.0) 
                    else x)
                )
        
        # Save the final file
        df.to_excel(output_file, index=False)
        
        # Clean up temporary file
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)
            
        print(f"Shopify feed saved to {output_file}")
    
    return shopify_feed

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate Shopify product feed from MASTER COPY Excel file')
    parser.add_argument('--input', '-i', default='MASTER COPY.xlsx', help='Input Excel file path')
    parser.add_argument('--output', '-o', help='Output Excel file path')
    parser.add_argument('--test', '-t', action='store_true', help='Run in test mode with example rows')
    parser.add_argument('--rows', '-r', help='Custom row numbers to process in format "start-end" (e.g., "14786-14787")')
    
    args = parser.parse_args()
    
    # If no output file specified, create one with timestamp
    if not args.output:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output = f'shopify_feed_{timestamp}.xlsx'
    
    # Handle custom row specification
    if args.rows:
        try:
            # Parse the row range
            start_row, end_row = map(int, args.rows.split('-'))
            print(f"Processing custom rows {start_row} to {end_row}")
            
            # Update the configuration
            CONFIG["test_start_row"] = start_row
            CONFIG["test_end_row"] = end_row
            
            # Load the Excel file
            df = pd.read_excel(args.input, sheet_name='MASTER COPY')
            df.index = df.index + 1  # Convert to 1-based indexing
            
            # Check if rows exist
            if start_row not in df.index or end_row not in df.index:
                print(f"Error: Row {start_row} or {end_row} not found in the MASTER COPY sheet")
                exit(1)
            
            # Always include the header row if not already included
            if start_row > 1:
                header_row = start_row - 1
                print(f"Including header row {header_row} for reference")
                
                # Extract the product rows including header row
                product_rows = df.loc[header_row:end_row]
            else:
                product_rows = df.loc[start_row:end_row]
                
            # Check if we have valid data
            if product_rows.empty:
                print("Error: No valid product rows found in the specified range")
                exit(1)
                
            # Find the first row with a description
            valid_rows = product_rows[~pd.isna(product_rows['description'])]
            
            if valid_rows.empty:
                print("Error: No rows with product description found in the specified range")
                exit(1)
                
            print(f"Found {len(product_rows)} rows for product: {valid_rows.iloc[0]['description']}")
            
            # Create custom test mode based on these rows
            args.test = True  # Enable test mode
        except ValueError:
            print("Error: Invalid row format. Please use format 'start-end' (e.g., '14786-14787')")
            exit(1)
        except Exception as e:
            print(f"Error processing custom rows: {e}")
            exit(1)
    
    feed = generate_shopify_feed(args.input, args.output, args.test)
    print(f"Generated {len(feed)} rows in the Shopify feed")
    
    # Print sample of the feed
    print("\nSample of generated Shopify feed:")
    sample_columns = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
    print(feed[sample_columns].head(10)) 