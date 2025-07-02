import pandas as pd
import numpy as np
import re
import argparse
from datetime import datetime
import warnings
import openpyxl

# Version information
__version__ = "1.10.0"
__date__ = "2025-01-15"
__description__ = "Shopify Product Feed Generator"

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

def clean_option_value(value, option_name):
    """Clean option values - specifically remove commas from Size options"""
    if pd.isna(value):
        return value
    
    # If this is a Size option, replace commas with ' -'
    if option_name == "Size":
        return str(value).replace(',', ' -')
    
    return value

def get_tags_from_column_k(product_group):
    """Get tags from column K (11th column, index 10) of the first row in product group"""
    try:
        # Get the first row of the product group
        first_row = product_group[0]
        
        # Handle pandas Series (both test and normal mode use Series)
        if hasattr(first_row, 'iloc'):  # This is a pandas Series
            if len(first_row) > 10:  # Make sure column K exists
                column_k_value = first_row.iloc[10]
                if not pd.isna(column_k_value) and str(column_k_value).strip():
                    return str(column_k_value).strip()
        
        return None
    except (IndexError, AttributeError, KeyError):
        return None

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
            # Look for exact matches or columns that start with the count followed by a space or parenthesis
            matching_count_cols = [col for col in finishes_df.columns 
                                   if (str(col) == str(finish_count) or 
                                       str(col).startswith(str(finish_count) + ' ') or
                                       str(col).startswith(str(finish_count) + '('))]
            if matching_count_cols:
                finish_col = matching_count_cols[0]
        except (ValueError, TypeError):
            pass
    
    # If still no match, check column F (6th column) and other specific finish code columns
    if not finish_col:
        # First check column F (which would be column 5 in 0-indexed or 6 in 1-indexed)
        # Try different ways column F might be represented
        f_column_candidates = [5, 6, 'F', 'f']
        for candidate in f_column_candidates:
            if candidate in finishes_df.columns:
                finishes_in_col = finishes_df[candidate].dropna().tolist()
                if len(finishes_in_col) > 0:
                    # Check if this column contains finish codes (entries with parentheses)
                    has_codes = any("(" in str(finish) and ")" in str(finish) for finish in finishes_in_col)
                    if has_codes:
                        finish_col = candidate
                        break
        
        # If column F doesn't work, look for other columns with specific finish codes
        if not finish_col:
            for col in finishes_df.columns:
                if col not in [0, '0', 25, '25']:  # Skip first and last default columns
                    finishes_in_col = finishes_df[col].dropna().tolist()
                    # Look for columns with a specific number of finishes that might be relevant
                    if len(finishes_in_col) > 0 and len(finishes_in_col) < 25:  # Less than the full set
                        # Check if this column contains finish codes (entries with parentheses)
                        has_codes = any("(" in str(finish) and ")" in str(finish) for finish in finishes_in_col)
                        if has_codes:
                            finish_col = col
                            break
    
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
    
    # Track products where finishes couldn't be identified
    finishes_not_found = []
    
    # Track products that couldn't be processed due to missing data
    products_not_processed = []
    
    # Create a mapping of finish codes to full names
    finish_code_to_name = {}
    for col in finishes_df.columns:
        finishes = finishes_df[col].dropna().tolist()
        for finish in finishes:
            if "(" in finish and ")" in finish:
                code = finish.split("(")[1].split(")")[0].strip()
                finish_code_to_name[code] = finish
    
    # Define which finishes belong to ## and x## categories
    hash_codes = ["PN", "SN", "BZ", "AB", "SB", "DB", "BAB", "BZW", "BABW", "ABW", "DBW", "NBW", "SBW", "PBUL"]
    xhash_codes = ["PCOP", "SCOP", "BLN", "PEW", "MBL", "ASV", "RGP", "ACOP"]
    
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
        
        # Filter out empty rows
        valid_rows = product_rows[~product_rows.isnull().all(axis=1)].copy()
        
        if valid_rows.empty:
            print("Error: No valid product rows found in the specified range")
            return pd.DataFrame()
        
        # Group products by description - this handles multiple products in the range
        product_groups = []
        current_group = []
        current_description = None
        
        for idx, row in valid_rows.iterrows():
            # If the row has a description and it's different from the current one, start a new group
            if not pd.isna(row['description']):
                if current_description != row['description']:
                    if current_group:
                        product_groups.append((current_description, current_group))
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
            product_groups.append((current_description, current_group))
        
        # Process each product group separately
        print(f"Found {len(product_groups)} distinct products in the row range")
        
        for product_num, (product_description, product_group_rows) in enumerate(product_groups):
            if pd.isna(product_description):
                print(f"Skipping product group {product_num+1} with no description")
                continue
                
            print(f"\nProcessing product {product_num+1}: {product_description}")
            
            # Get tags from column K for this product group
            tags = get_tags_from_column_k(product_group_rows)
            if not tags:
                print(f"Warning: No tags found in column K for product: {product_description}")
                # Track this product as not processed due to missing tags
                products_not_processed.append({
                    "Product Description": product_description,
                    "Reason": "Missing tag in column K",
                    "Row Range": f"Rows {product_group_rows[0].name}-{product_group_rows[-1].name}" if len(product_group_rows) > 0 else "Unknown"
                })
                continue
            else:
                print(f"Found tags for product: {tags}")
            
            # Generate handle from product description
            handle = clean_string(product_description)
            
            # Convert the list of rows back to a DataFrame
            valid_rows_df = pd.DataFrame(product_group_rows)
            
            # Get all unique sizes from the valid rows
            unique_sizes = valid_rows_df['size'].dropna().unique()
            print(f"Unique sizes: {unique_sizes}")
            print(f"Number of unique sizes: {len(unique_sizes)}")
            
            # Check if product name contains keywords for finish selection
            keywords = ['Bjorn', 'Cadiz', 'Denham', 'Wilton', 'Capri', 'Leon', 'Oxon']
            matching_keyword = None
            for keyword in keywords:
                if keyword.lower() in product_description.lower():
                    matching_keyword = keyword
                    break
                    
            # Find matching finish column based on product name
            product_specific_finishes = None
            if matching_keyword:
                for col in finishes_df.columns:
                    if matching_keyword.lower() in str(col).lower():
                        product_specific_finishes = finishes_df[col].dropna().tolist()
                        print(f"Found product-specific finishes for '{matching_keyword}': {len(product_specific_finishes)} finishes")
                        break
            
            # Check if finish count is specified in any row
            finish_count_specific_finishes = None
            for idx, row in valid_rows_df.iterrows():
                if not pd.isna(row.get('finish count')):
                    try:
                        finish_count = int(row['finish count'])
                        for col in finishes_df.columns:
                            if str(col) == str(finish_count) or str(col).startswith(str(finish_count) + ' '):
                                finish_count_specific_finishes = finishes_df[col].dropna().tolist()
                                print(f"Found finish count {finish_count} with {len(finish_count_specific_finishes)} finishes")
                                break
                        if finish_count_specific_finishes:
                            break
                    except (ValueError, TypeError):
                        pass
            
            # Store SKU/price data by row and track which finishes each row applies to
            row_data = {}
            for idx, row in valid_rows_df.iterrows():
                if not pd.isna(row['code']) and not pd.isna(row['rrp']):
                    size = row['size'] if not pd.isna(row.get('size')) else None
                    # Convert SKU to string and remove .0 if it's a whole number
                    sku = str(row['code'])
                    if sku.endswith('.0'):
                        sku = sku[:-2]
                    price = float(row['rrp'])
                    finish_code = row['finish'] if not pd.isna(row['finish']) else None
                    
                    # Determine which finishes this row applies to
                    applicable_finishes = []
                    
                    # First priority: Use product-specific finishes if available
                    if product_specific_finishes:
                        applicable_finishes = product_specific_finishes
                        print(f"Row {idx}: Using {len(applicable_finishes)} product-specific finishes for '{matching_keyword}'")
                    
                    # Second priority: Use finish count specific finishes if available
                    elif finish_count_specific_finishes:
                        applicable_finishes = finish_count_specific_finishes
                        print(f"Row {idx}: Using {len(applicable_finishes)} finishes based on finish count")
                    
                    # Third priority: Use finish code
                    elif finish_code == "##":
                        # This row applies to the 14 ## finishes
                        applicable_finishes = [finish_code_to_name.get(code, f"Unknown ({code})") for code in hash_codes if code in finish_code_to_name]
                        if size:
                            print(f"Row {idx}: Size={size}, SKU={sku}, Price=£{price}, Finish=##, Applies to {len(applicable_finishes)} finishes")
                        else:
                            print(f"Row {idx}: No size, SKU={sku}, Price=£{price}, Finish=##, Applies to {len(applicable_finishes)} finishes")
                    elif finish_code == "x##":
                        # This row applies to the 8 x## finishes
                        applicable_finishes = [finish_code_to_name.get(code, f"Unknown ({code})") for code in xhash_codes if code in finish_code_to_name]
                        if size:
                            print(f"Row {idx}: Size={size}, SKU={sku}, Price=£{price}, Finish=x##, Applies to {len(applicable_finishes)} finishes")
                        else:
                            print(f"Row {idx}: No size, SKU={sku}, Price=£{price}, Finish=x##, Applies to {len(applicable_finishes)} finishes")
                    elif finish_code in finish_code_to_name:
                        # This row applies to a specific finish
                        applicable_finishes = [finish_code_to_name[finish_code]]
                        if size:
                            print(f"Row {idx}: Size={size}, SKU={sku}, Price=£{price}, Finish={finish_code}, Applies to {finish_code_to_name[finish_code]}")
                        else:
                            print(f"Row {idx}: No size, SKU={sku}, Price=£{price}, Finish={finish_code}, Applies to {finish_code_to_name[finish_code]}")
                    else:
                        # If we can't determine the finishes, use all finishes from column 25
                        print(f"Warning: Row {idx} has unknown finish code {finish_code}. Using all finishes.")
                        applicable_finishes = finishes_df[25].dropna().tolist()
                        
                        # Track this product as having unidentified finishes
                        finishes_not_found.append({
                            "Product Description": product_description,
                            "Row Index": idx,
                            "Size": size if size else "No size",
                            "SKU": sku,
                            "Finish Code": finish_code,
                            "Reason": f"Unknown finish code: {finish_code}",
                            "Defaulted To": f"{len(applicable_finishes)} finishes from column 25"
                        })
                    
                    row_data[idx] = {
                        "size": size,
                        "price": price,
                        "sku": sku,
                        "finish_code": finish_code,
                        "applicable_finishes": applicable_finishes,
                        "has_size": size is not None
                    }
            
            if not row_data:
                print(f"Error: No valid SKU/price data found in rows for product: {product_description}")
                # Track this product as not processed
                products_not_processed.append({
                    "Product Description": product_description,
                    "Reason": "Missing SKU/price data",
                    "Row Range": f"Rows {product_group_rows[0].name}-{product_group_rows[-1].name}" if len(product_group_rows) > 0 else "Unknown"
                })
                continue
            
            print(f"Found {len(row_data)} rows with valid data")
            
            # Get all unique finishes that will be used
            all_applicable_finishes = []
            for data in row_data.values():
                all_applicable_finishes.extend(data["applicable_finishes"])
            unique_finishes = list(set(all_applicable_finishes))
            print(f"Total unique finishes: {len(unique_finishes)}")
            
            # Determine if this product has sizes or not  
            product_has_sizes = len(unique_sizes) > 0
            
            # Calculate expected number of variants
            if product_has_sizes:
                expected_variants = len(unique_sizes) * len(unique_finishes)
                print(f"Expected number of variants: {len(unique_sizes)} sizes × {len(unique_finishes)} finishes = {expected_variants}")
            else:
                expected_variants = len(unique_finishes)
                print(f"Expected number of variants (no sizes): {len(unique_finishes)} finishes = {expected_variants}")
            
            # Create rows for each size-finish combination, but only where the SKU applies to that finish
            product_rows = []
            
            # Track if we've already set the first-row-only fields
            first_row_set = False
            
            if product_has_sizes:
                # Process each size
                for size in unique_sizes:
                    # For each finish, create a row in the Shopify feed
                    for finish in unique_finishes:
                        # Find which row's SKU applies to this specific finish
                        matching_row = None
                        
                        # First try to find an exact match for this specific finish
                        for idx, data in row_data.items():
                            if data["size"] == size:
                                # Check if this is a specific finish code (not ## or x##)
                                if data["finish_code"] in finish_code_to_name and finish == finish_code_to_name[data["finish_code"]]:
                                    matching_row = (idx, data)
                                    break
                        
                        # If no exact match found, check for ## and x## categories
                        if matching_row is None:
                            # Check if this finish is in the hash_codes (##) list
                            finish_code = None
                            for code in hash_codes:
                                if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                    finish_code = "##"
                                    break
                            
                            # If not in hash_codes, check if it's in xhash_codes (x##)
                            if finish_code is None:
                                for code in xhash_codes:
                                    if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                        finish_code = "X##"
                                        break
                            
                            # Now find the row with this finish_code
                            if finish_code:
                                for idx, data in row_data.items():
                                    if data["size"] == size and data["finish_code"] == finish_code:
                                        matching_row = (idx, data)
                                        break
                        
                        # If still no match, fall back to the original method as a last resort
                        if matching_row is None:
                            for idx, data in row_data.items():
                                if data["size"] == size and finish in data["applicable_finishes"]:
                                    matching_row = (idx, data)
                                    break
                        
                        # If no row applies to this finish, skip it
                        if matching_row is None:
                            continue
                        
                        idx, data = matching_row
                        sku_base = data["sku"]
                        price = data["price"]
                        
                        # Create new row based on template
                        new_row = {col: None for col in template_columns}
                        
                        # Set values based on mapping instructions
                        new_row['Handle'] = handle
                        
                        # Only set certain fields for the first row of the product
                        is_first_row = not first_row_set
                        
                        if is_first_row:
                            new_row['Title'] = product_description
                            # Set Image Alt Text to match the Title
                            new_row['Image Alt Text'] = product_description
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
                            
                            # Set Tags from column K
                            new_row['Tags'] = tags
                            
                            # Set regional inclusion - use "TRUE" string instead of boolean True
                            for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                                new_row[f'Included / {region}'] = "TRUE"
                            
                            new_row['Status'] = "draft"
                            
                            # Mark that we've set the first row fields
                            first_row_set = True
                        
                        # Set variant-specific values
                        new_row['Option1 Value'] = clean_option_value(size, "Size")
                        new_row['Option2 Value'] = finish
                        
                        # Set SKU - the original SKU from the row that applies to this finish
                        new_row['Variant SKU'] = sku_base
                        new_row['Variant Grams'] = 0
                        new_row['Variant Inventory Tracker'] = "shopify"
                        new_row['Variant Inventory Qty'] = 10000
                        new_row['Variant Inventory Policy'] = "deny"
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
            else:
                # Process products without sizes - finishes become Option1
                for finish in unique_finishes:
                    # Find which row's SKU applies to this specific finish
                    matching_row = None
                    
                    # First try to find an exact match for this specific finish
                    for idx, data in row_data.items():
                        if not data["has_size"]:  # Only consider rows without sizes
                            # Check if this is a specific finish code (not ## or x##)
                            if data["finish_code"] in finish_code_to_name and finish == finish_code_to_name[data["finish_code"]]:
                                matching_row = (idx, data)
                                break
                    
                    # If no exact match found, check for ## and x## categories
                    if matching_row is None:
                        # Check if this finish is in the hash_codes (##) list
                        finish_code = None
                        for code in hash_codes:
                            if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                finish_code = "##"
                                break
                        
                        # If not in hash_codes, check if it's in xhash_codes (x##)
                        if finish_code is None:
                            for code in xhash_codes:
                                if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                    finish_code = "X##"
                                    break
                        
                        # Now find the row with this finish_code
                        if finish_code:
                            for idx, data in row_data.items():
                                if not data["has_size"] and data["finish_code"] == finish_code:
                                    matching_row = (idx, data)
                                    break
                    
                    # If still no match, fall back to the original method as a last resort
                    if matching_row is None:
                        for idx, data in row_data.items():
                            if not data["has_size"] and finish in data["applicable_finishes"]:
                                matching_row = (idx, data)
                                break
                    
                    # If no row applies to this finish, skip it
                    if matching_row is None:
                        continue
                    
                    idx, data = matching_row
                    sku_base = data["sku"]
                    price = data["price"]
                    
                    # Create new row based on template
                    new_row = {col: None for col in template_columns}
                    
                    # Set values based on mapping instructions
                    new_row['Handle'] = handle
                    
                    # Only set certain fields for the first row of the product
                    is_first_row = not first_row_set
                    
                    if is_first_row:
                        new_row['Title'] = product_description
                        # Set Image Alt Text to match the Title
                        new_row['Image Alt Text'] = product_description
                        new_row['Vendor'] = "vendor-unknown"
                        new_row['Product Category'] = "Uncategorized"
                        new_row['Type'] = get_product_type(product_description)
                        # Use string "TRUE" instead of boolean True
                        new_row['Published'] = "TRUE"
                        
                        # For products without sizes - set Option1 to Finish, leave Option2 empty
                        new_row['Option1 Name'] = "Finish"
                        # Option2 Name is left as None/empty
                        
                        # Use an example image from sample
                        if not sample_df.empty and 'Image Src' in sample_df.columns and not pd.isna(sample_df['Image Src'].iloc[0]):
                            new_row['Image Src'] = sample_df['Image Src'].iloc[0]
                            
                        new_row['Image Position'] = 1
                        new_row['Gift Card'] = "FALSE"
                        new_row['SEO Title'] = f"{product_description} | A&H Brass"
                        
                        # Set Tags from column K
                        new_row['Tags'] = tags
                        
                        # Set regional inclusion - use "TRUE" string instead of boolean True
                        for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                            new_row[f'Included / {region}'] = "TRUE"
                        
                        new_row['Status'] = "draft"
                        
                        # Mark that we've set the first row fields
                        first_row_set = True
                    
                    # Set variant-specific values for products without sizes
                    new_row['Option1 Value'] = clean_option_value(finish, "Finish")  # Finish goes to Option1
                    # Option2 Value is left as None/empty
                    
                    # Set SKU - the original SKU from the row that applies to this finish
                    new_row['Variant SKU'] = sku_base
                    new_row['Variant Grams'] = 0
                    new_row['Variant Inventory Tracker'] = "shopify"
                    new_row['Variant Inventory Qty'] = 10000
                    new_row['Variant Inventory Policy'] = "deny"
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
            
            # Get tags from column K for this product group
            tags = get_tags_from_column_k(product_group)
            if not tags:
                print(f"Warning: No tags found in column K for product: {product_description}")
                # Track this product as not processed due to missing tags
                products_not_processed.append({
                    "Product Description": product_description,
                    "Reason": "Missing tag in column K",
                    "Row Range": f"Product group with {len(product_group)} rows"
                })
                continue
            else:
                print(f"Found tags for product: {tags}")
            
            # Generate handle from product description
            handle = clean_string(product_description)
            
            # Determine product type
            product_type = get_product_type(product_description)
            
            # Get unique sizes for this product, but keep all rows for processing
            all_rows = product_group
            valid_rows = [row for row in product_group if not pd.isna(row.get('code')) and not pd.isna(row.get('rrp'))]
            
            # Check if this product has any sizes
            rows_with_sizes = [row for row in valid_rows if not pd.isna(row.get('size'))]
            rows_without_sizes = [row for row in valid_rows if pd.isna(row.get('size'))]
            
            has_sizes = len(rows_with_sizes) > 0
            unique_sizes = list(set([row['size'] for row in rows_with_sizes if not pd.isna(row['size'])])) if has_sizes else []
            
            print(f"  Found {len(unique_sizes)} unique sizes, {len(rows_without_sizes)} rows without sizes")
            
            # Use all valid rows (both with and without sizes) for processing
            valid_rows = rows_with_sizes + rows_without_sizes
            
            # Check if product name contains keywords for finish selection
            keywords = ['Bjorn', 'Cadiz', 'Denham', 'Wilton', 'Capri', 'Leon', 'Oxon']
            matching_keyword = None
            for keyword in keywords:
                if keyword.lower() in product_description.lower():
                    matching_keyword = keyword
                    break
                    
            # Find matching finish column based on product name
            product_specific_finishes = None
            if matching_keyword:
                for col in finishes_df.columns:
                    if matching_keyword.lower() in str(col).lower():
                        product_specific_finishes = finishes_df[col].dropna().tolist()
                        print(f"  Found product-specific finishes for '{matching_keyword}': {len(product_specific_finishes)} finishes")
                        break
            
            # Check if finish count is specified in any row
            finish_count_specific_finishes = None
            for i, row in enumerate(valid_rows):
                if not pd.isna(row.get('finish count')):
                    try:
                        finish_count = int(row['finish count'])
                        for col in finishes_df.columns:
                            if str(col) == str(finish_count) or str(col).startswith(str(finish_count) + ' '):
                                finish_count_specific_finishes = finishes_df[col].dropna().tolist()
                                print(f"  Found finish count {finish_count} with {len(finish_count_specific_finishes)} finishes")
                                break
                        if finish_count_specific_finishes:
                            break
                    except (ValueError, TypeError):
                        pass
            
            # Store data by row and track which finishes each row applies to
            row_data = {}
            for i, row in enumerate(valid_rows):
                if pd.isna(row.get('code')) or pd.isna(row.get('rrp')):
                    continue
                    
                size = row['size'] if not pd.isna(row.get('size')) else None
                # Convert SKU to string and remove .0 if it's a whole number
                sku = str(row['code'])
                if sku.endswith('.0'):
                    sku = sku[:-2]
                price = float(row['rrp'])
                finish_code = row['finish'] if not pd.isna(row['finish']) else None
                
                # Determine which finishes this row applies to
                applicable_finishes = []
                
                # First priority: Use product-specific finishes if available
                if product_specific_finishes:
                    applicable_finishes = product_specific_finishes
                    print(f"  Row {i}: Using {len(applicable_finishes)} product-specific finishes for '{matching_keyword}'")
                
                # Second priority: Use finish count specific finishes if available
                elif finish_count_specific_finishes:
                    applicable_finishes = finish_count_specific_finishes
                    print(f"  Row {i}: Using {len(applicable_finishes)} finishes based on finish count")
                
                # Third priority: Use finish code
                elif finish_code == "##":
                    # This row applies to the 14 ## finishes
                    applicable_finishes = [finish_code_to_name.get(code, f"Unknown ({code})") for code in hash_codes if code in finish_code_to_name]
                    print(f"  Row {i}: Size={size}, SKU={sku}, Price=£{price}, Finish=##, Applies to {len(applicable_finishes)} finishes")
                elif finish_code == "x##":
                    # This row applies to the 8 x## finishes
                    applicable_finishes = [finish_code_to_name.get(code, f"Unknown ({code})") for code in xhash_codes if code in finish_code_to_name]
                    print(f"  Row {i}: Size={size}, SKU={sku}, Price=£{price}, Finish=x##, Applies to {len(applicable_finishes)} finishes")
                elif finish_code in finish_code_to_name:
                    # This row applies to a specific finish
                    applicable_finishes = [finish_code_to_name[finish_code]]
                    print(f"  Row {i}: Size={size}, SKU={sku}, Price=£{price}, Finish={finish_code}, Applies to {finish_code_to_name[finish_code]}")
                else:
                    # If we can't determine the finishes, use all finishes from column 25
                    print(f"  Warning: Row {i} has unknown finish code {finish_code}. Using all finishes.")
                    applicable_finishes = finishes_df[25].dropna().tolist()
                    
                    # Track this product as having unidentified finishes
                    finishes_not_found.append({
                        "Product Description": product_description,
                        "Row Index": f"Row {i}",
                        "Size": size,
                        "SKU": sku,
                        "Finish Code": finish_code,
                        "Reason": f"Unknown finish code: {finish_code}",
                        "Defaulted To": f"{len(applicable_finishes)} finishes from column 25"
                    })
                
                row_data[i] = {
                    "size": size,
                    "price": price,
                    "sku": sku,
                    "finish_code": finish_code,
                    "applicable_finishes": applicable_finishes,
                    "row": row,  # Keep original row data for reference
                    "has_size": size is not None
                }
            
            if not row_data:
                print(f"Error: No valid SKU/price data found in rows for product: {product_description}")
                # Track this product as not processed
                products_not_processed.append({
                    "Product Description": product_description,
                    "Reason": "Missing SKU/price data",
                    "Row Range": f"Rows {product_group_rows[0].name}-{product_group_rows[-1].name}" if len(product_group_rows) > 0 else "Unknown"
                })
                continue
            
            # Get all unique finishes that will be used
            all_applicable_finishes = []
            for data in row_data.values():
                all_applicable_finishes.extend(data["applicable_finishes"])
            unique_finishes = list(set(all_applicable_finishes))
            
            # Determine if this product has sizes or not
            product_has_sizes = len(unique_sizes) > 0
            
            # Calculate expected number of variants
            if product_has_sizes:
                expected_variants = len(unique_sizes) * len(unique_finishes)
                print(f"  Expected variants: {len(unique_sizes)} sizes × {len(unique_finishes)} finishes = {expected_variants}")
            else:
                expected_variants = len(unique_finishes)
                print(f"  Expected variants (no sizes): {len(unique_finishes)} finishes = {expected_variants}")
            
            # Create Shopify rows for each size-finish combination, but only where the SKU applies
            is_first_row = True
            
            if product_has_sizes:
                # Process each size
                for size in unique_sizes:
                    # For each finish, create a row in the Shopify feed
                    for finish in unique_finishes:
                        # Find which row's SKU applies to this specific finish
                        matching_row = None
                        
                        # First try to find an exact match for this specific finish
                        for idx, data in row_data.items():
                            if data["size"] == size:
                                # Check if this is a specific finish code (not ## or x##)
                                if data["finish_code"] in finish_code_to_name and finish == finish_code_to_name[data["finish_code"]]:
                                    matching_row = (idx, data)
                                    break
                        
                        # If no exact match found, check for ## and x## categories
                        if matching_row is None:
                            # Check if this finish is in the hash_codes (##) list
                            finish_code = None
                            for code in hash_codes:
                                if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                    finish_code = "##"
                                    break
                            
                            # If not in hash_codes, check if it's in xhash_codes (x##)
                            if finish_code is None:
                                for code in xhash_codes:
                                    if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                        finish_code = "X##"
                                        break
                            
                            # Now find the row with this finish_code
                            if finish_code:
                                for idx, data in row_data.items():
                                    if data["size"] == size and data["finish_code"] == finish_code:
                                        matching_row = (idx, data)
                                        break
                        
                        # If still no match, fall back to the original method as a last resort
                        if matching_row is None:
                            for idx, data in row_data.items():
                                if data["size"] == size and finish in data["applicable_finishes"]:
                                    matching_row = (idx, data)
                                    break
                        
                        # If no row applies to this finish, skip it
                        if matching_row is None:
                            continue
                        
                        idx, data = matching_row
                        sku_base = data["sku"]
                        price = data["price"]
                        
                        # Create new row based on template
                        new_row = {col: None for col in template_columns}
                        
                        # Set values based on mapping instructions
                        new_row['Handle'] = handle
                        
                        # Only set certain fields for the first row of the product
                        if is_first_row:
                            new_row['Title'] = product_description
                            # Set Image Alt Text to match the Title
                            new_row['Image Alt Text'] = product_description
                            new_row['Vendor'] = "vendor-unknown"
                            new_row['Product Category'] = "Uncategorized"
                            new_row['Type'] = product_type
                            # Use string "TRUE" instead of boolean True
                            new_row['Published'] = "TRUE"
                            
                            # For products with sizes - set Option1 to Size, Option2 to Finish
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
                            
                            # Set Tags from column K
                            new_row['Tags'] = tags
                            
                            # Set regional inclusion - use "TRUE" string instead of boolean True
                            for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                                new_row[f'Included / {region}'] = "TRUE"
                            
                            new_row['Status'] = "draft"
                            
                            # Mark that we've set the first row fields
                            is_first_row = False
                        
                        # Set variant-specific values for products with sizes
                        new_row['Option1 Value'] = clean_option_value(size, "Size")
                        new_row['Option2 Value'] = finish
                        new_row['Variant SKU'] = sku_base
                        new_row['Variant Grams'] = 0
                        new_row['Variant Inventory Tracker'] = "shopify"
                        new_row['Variant Inventory Qty'] = 10000
                        new_row['Variant Inventory Policy'] = "deny"
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
            else:
                # Process products without sizes - finishes become Option1
                for finish in unique_finishes:
                    # Find which row's SKU applies to this specific finish
                    matching_row = None
                    
                    # First try to find an exact match for this specific finish
                    for idx, data in row_data.items():
                        if not data["has_size"]:  # Only consider rows without sizes
                            # Check if this is a specific finish code (not ## or x##)
                            if data["finish_code"] in finish_code_to_name and finish == finish_code_to_name[data["finish_code"]]:
                                matching_row = (idx, data)
                                break
                    
                    # If no exact match found, check for ## and x## categories
                    if matching_row is None:
                        # Check if this finish is in the hash_codes (##) list
                        finish_code = None
                        for code in hash_codes:
                            if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                finish_code = "##"
                                break
                        
                        # If not in hash_codes, check if it's in xhash_codes (x##)
                        if finish_code is None:
                            for code in xhash_codes:
                                if code in finish_code_to_name and finish == finish_code_to_name[code]:
                                    finish_code = "X##"
                                    break
                        
                        # Now find the row with this finish_code
                        if finish_code:
                            for idx, data in row_data.items():
                                if not data["has_size"] and data["finish_code"] == finish_code:
                                    matching_row = (idx, data)
                                    break
                    
                    # If still no match, fall back to the original method as a last resort
                    if matching_row is None:
                        for idx, data in row_data.items():
                            if not data["has_size"] and finish in data["applicable_finishes"]:
                                matching_row = (idx, data)
                                break
                    
                    # If no row applies to this finish, skip it
                    if matching_row is None:
                        continue
                    
                    idx, data = matching_row
                    sku_base = data["sku"]
                    price = data["price"]
                    
                    # Create new row based on template
                    new_row = {col: None for col in template_columns}
                    
                    # Set values based on mapping instructions
                    new_row['Handle'] = handle
                    
                    # Only set certain fields for the first row of the product
                    if is_first_row:
                        new_row['Title'] = product_description
                        # Set Image Alt Text to match the Title
                        new_row['Image Alt Text'] = product_description
                        new_row['Vendor'] = "vendor-unknown"
                        new_row['Product Category'] = "Uncategorized"
                        new_row['Type'] = product_type
                        # Use string "TRUE" instead of boolean True
                        new_row['Published'] = "TRUE"
                        
                        # For products without sizes - set Option1 to Finish, leave Option2 empty
                        new_row['Option1 Name'] = "Finish"
                        # Option2 Name is left as None/empty
                        
                        # Use an example image from sample
                        if not sample_df.empty and 'Image Src' in sample_df.columns and not pd.isna(sample_df['Image Src'].iloc[0]):
                            new_row['Image Src'] = sample_df['Image Src'].iloc[0]
                            
                        new_row['Image Position'] = 1
                        new_row['Gift Card'] = "FALSE"
                        new_row['SEO Title'] = f"{product_description} | A&H Brass"
                        
                        # Set Tags from column K
                        new_row['Tags'] = tags
                        
                        # Set regional inclusion - use "TRUE" string instead of boolean True
                        for region in ['United Kingdom', 'Australia', 'Canada', 'Europe', 'International', 'United States']:
                            new_row[f'Included / {region}'] = "TRUE"
                        
                        new_row['Status'] = "draft"
                        
                        # Mark that we've set the first row fields
                        is_first_row = False
                    
                    # Set variant-specific values for products without sizes
                    new_row['Option1 Value'] = clean_option_value(finish, "Finish")  # Finish goes to Option1
                    # Option2 Value is left as None/empty
                    new_row['Variant SKU'] = sku_base
                    new_row['Variant Grams'] = 0
                    new_row['Variant Inventory Tracker'] = "shopify"
                    new_row['Variant Inventory Qty'] = 10000
                    new_row['Variant Inventory Policy'] = "deny"
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
    
    # Export finishes not found to CSV if there are any
    if finishes_not_found:
        finishes_not_found_df = pd.DataFrame(finishes_not_found)
        csv_filename = "finishes_not_found.csv"
        finishes_not_found_df.to_csv(csv_filename, index=False)
        print(f"⚠️  Found {len(finishes_not_found)} products with unidentified finishes")
        print(f"📄 Details exported to {csv_filename}")
    else:
        print("✅ All products had identifiable finishes")
    
    # Export products not processed to CSV if there are any
    if products_not_processed:
        products_not_processed_df = pd.DataFrame(products_not_processed)
        csv_filename = "products_not_processed.csv"
        products_not_processed_df.to_csv(csv_filename, index=False)
        print(f"⚠️  Found {len(products_not_processed)} products that couldn't be processed")
        print(f"📄 Details exported to {csv_filename}")
    
    return shopify_feed, finishes_not_found, products_not_processed

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate Shopify product feed from MASTER COPY Excel file')
    parser.add_argument('--input', '-i', default='MASTER COPY.xlsx', help='Input Excel file path')
    parser.add_argument('--output', '-o', help='Output Excel file path')
    parser.add_argument('--test', '-t', action='store_true', help='Run in test mode with example rows')
    parser.add_argument('--rows', '-r', help='Custom row numbers to process in format "start-end" (e.g., "14786-14787")')
    parser.add_argument('--version', '-v', action='store_true', help='Display version information')
    
    args = parser.parse_args()
    
    # Display version information if requested
    if args.version:
        print(f"{__description__} v{__version__} ({__date__})")
        exit(0)
        
    # Print version header
    print(f"Running {__description__} v{__version__}")
    
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
            
            # Check if rows exist in the Excel file
            # Use openpyxl to get accurate row count (pandas can miss rows)
            try:
                wb = openpyxl.load_workbook(args.input, read_only=True)
                sheet = wb['MASTER COPY']
                max_row = sheet.max_row
                wb.close()
                
                if start_row > max_row:
                    print(f"Error: Row {start_row} not found in the MASTER COPY sheet (max row is {max_row})")
                    exit(1)
                
                # If end_row is beyond the last row, adjust it
                if end_row > max_row:
                    print(f"Warning: Row {end_row} exceeds max row {max_row} in the MASTER COPY sheet. Adjusting to last available row.")
                    end_row = max_row
                    # Update the configuration
                    CONFIG["test_end_row"] = end_row
            except Exception as e:
                print(f"Warning: Unable to verify row range with openpyxl: {e}")
                print("Falling back to pandas row verification...")
                
                # Original pandas check as fallback
                if start_row not in df.index:
                    print(f"Error: Row {start_row} not found in the MASTER COPY sheet (max row is {df.index[-1]})")
                    exit(1)
                
                if end_row not in df.index:
                    last_row = df.index[-1]
                    print(f"Warning: Row {end_row} not found in the MASTER COPY sheet. Adjusting to last available row: {last_row}")
                    end_row = last_row
                    # Update the configuration
                    CONFIG["test_end_row"] = end_row
            
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
    
    feed, finishes_not_found, products_not_processed = generate_shopify_feed(args.input, args.output, args.test)
    print(f"Generated {len(feed)} rows in the Shopify feed")
    
    # Print sample of the feed
    print("\nSample of generated Shopify feed:")
    sample_columns = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
    print(feed[sample_columns].head(10)) 