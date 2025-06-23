import pandas as pd
import openpyxl

# First check the raw data to understand what's in rows 14786-14787
print("Examining raw data for rows 14786-14787:")
try:
    # Check with openpyxl first (most accurate row detection)
    wb = openpyxl.load_workbook('MASTER COPY.xlsx', read_only=True)
    sheet = wb['MASTER COPY']
    
    # Get all values from rows 14786-14787
    for row_num in [14786, 14787]:
        print(f"\nRow {row_num} data:")
        row_values = {}
        for col in range(1, 20):  # Check first 20 columns
            cell_value = sheet.cell(row=row_num, column=col).value
            if cell_value is not None:
                row_values[col] = cell_value
        print(row_values)
    
    wb.close()
    
    # Now check with pandas to get column names
    df = pd.read_excel('MASTER COPY.xlsx', sheet_name='MASTER COPY')
    df.index = df.index + 1  # Convert to 1-based indexing
    
    # Get specific values we're interested in
    for row_num in [14786, 14787]:
        if row_num in df.index:
            row = df.loc[row_num]
            print(f"\nRow {row_num} with column names:")
            for col, value in row.items():
                if not pd.isna(value):
                    print(f"  {col}: {value}")
    
    # Check the finishes available in the Finishes tab
    finishes_df = pd.read_excel('MASTER COPY.xlsx', sheet_name='Finishes')
    print("\nColumns in Finishes tab:")
    print(finishes_df.columns.tolist())
    
    # Check for Cadiz specifically
    cadiz_col = None
    for col in finishes_df.columns:
        if 'cadiz' in str(col).lower():
            cadiz_col = col
            break
    
    if cadiz_col:
        print(f"\nFinishes in Cadiz column ({cadiz_col}):")
        cadiz_finishes = finishes_df[cadiz_col].dropna().tolist()
        print(f"Found {len(cadiz_finishes)} Cadiz finishes:")
        for i, finish in enumerate(cadiz_finishes):
            print(f"  {i+1}. {finish}")
    
    # Check column 8 (potential finish count column)
    if 8 in finishes_df.columns:
        print(f"\nFinishes in column 8:")
        finishes_8 = finishes_df[8].dropna().tolist()
        print(f"Found {len(finishes_8)} finishes in column 8:")
        for i, finish in enumerate(finishes_8):
            print(f"  {i+1}. {finish}")
    
    # Now run the script on these rows to see what finishes it's using
    print("\n\nRunning the script on rows 14786-14787 to see which finishes it uses...")
    from shopify_feed_generator import CONFIG
    
    # Create a test function to mimic the script's behavior but only extract finish info
    def check_finish_detection(rows):
        """Check which finishes the script would use for these rows"""
        df = pd.read_excel('MASTER COPY.xlsx', sheet_name='MASTER COPY')
        df.index = df.index + 1  # Convert to 1-based indexing
        finishes_df = pd.read_excel('MASTER COPY.xlsx', sheet_name='Finishes')
        
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
        
        # Get product details
        product_rows = df.loc[rows[0]:rows[1]].copy()
        valid_rows = product_rows[~pd.isna(product_rows['description'])].copy()
        
        if valid_rows.empty:
            print("No valid rows found")
            return
        
        # Get product description
        product_description = valid_rows.iloc[0]['description']
        print(f"Product: {product_description}")
        
        # Process each row to determine which finishes it applies to
        for idx, row in valid_rows.iterrows():
            if pd.isna(row.get('size')) or pd.isna(row.get('code')) or pd.isna(row.get('rrp')):
                continue
                
            size = row['size']
            sku = str(row['code'])
            price = float(row['rrp'])
            finish_code = row['finish'] if not pd.isna(row['finish']) else None
            finish_count = row['finish count'] if not pd.isna(row['finish count']) else None
            
            print(f"\nRow {idx}: Size={size}, SKU={sku}, Price={price}, Finish={finish_code}, Finish Count={finish_count}")
            
            # Check if product name contains keywords for finish selection
            keywords = ['Bjorn', 'Cadiz', 'Denham', 'Wilton', 'Capri', 'Leon', 'Oxon']
            matching_keyword = None
            for keyword in keywords:
                if keyword.lower() in product_description.lower():
                    matching_keyword = keyword
                    break
                    
            if matching_keyword:
                print(f"  Found keyword '{matching_keyword}' in product description")
                
                # Look for a matching column in Finishes tab
                matching_col = None
                for col in finishes_df.columns:
                    if matching_keyword.lower() in str(col).lower():
                        matching_col = col
                        break
                
                if matching_col:
                    matching_finishes = finishes_df[matching_col].dropna().tolist()
                    print(f"  Found matching column '{matching_col}' with {len(matching_finishes)} finishes")
                    print(f"  Would use these finishes: {matching_finishes}")
                else:
                    print(f"  No matching column found for '{matching_keyword}'")
            
            # Check if finish count is specified
            if finish_count is not None:
                print(f"  Finish count is specified: {finish_count}")
                
                # Look for a column matching this count
                matching_count_col = None
                for col in finishes_df.columns:
                    if str(col) == str(finish_count) or str(col).startswith(str(finish_count) + ' '):
                        matching_count_col = col
                        break
                
                if matching_count_col:
                    count_finishes = finishes_df[matching_count_col].dropna().tolist()
                    print(f"  Found matching count column '{matching_count_col}' with {len(count_finishes)} finishes")
                    print(f"  Would use these finishes: {count_finishes}")
                else:
                    print(f"  No matching column found for count {finish_count}")
            
            # Determine which finishes this row applies to based on the current logic
            applicable_finishes = []
            
            if finish_code == "##":
                # This row applies to the 14 ## finishes
                applicable_finishes = [finish_code_to_name.get(code, f"Unknown ({code})") for code in hash_codes if code in finish_code_to_name]
                print(f"  ## finishes: Would use {len(applicable_finishes)} finishes")
            elif finish_code == "x##":
                # This row applies to the 8 x## finishes
                applicable_finishes = [finish_code_to_name.get(code, f"Unknown ({code})") for code in xhash_codes if code in finish_code_to_name]
                print(f"  x## finishes: Would use {len(applicable_finishes)} finishes")
            elif finish_code in finish_code_to_name:
                # This row applies to a specific finish
                applicable_finishes = [finish_code_to_name[finish_code]]
                print(f"  Specific finish: Would use {finish_code_to_name[finish_code]}")
    
    # Run the test function
    check_finish_detection([14786, 14787])
    
except Exception as e:
    print(f"Error: {e}") 