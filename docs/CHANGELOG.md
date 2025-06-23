# Changelog for shopify_feed_generator.py

## Version 1.7.1 - 2023-06-05 (Streamlit App Enhancements)

### Improvements:
- Fixed Excel row detection in Streamlit app using openpyxl for accurate counting
- Now correctly detects all rows including row 14812
- Enhanced bar chart visualization with better colors and more compact design
- Improved UI layout with better spacing and organization
- Added row count information to the Excel file preview

## Version 1.7.0 - 2023-06-05 (Streamlit App)

### Improvements:
- Added a user-friendly Streamlit web application
- Allows for visual selection of row ranges
- Preview of products before generating feed
- Interactive product distribution visualization
- Easy download of generated feed
- Added proper documentation for the app

## Version 1.6.0 - 2023-06-05 (Multiple Product Support)

### Improvements:
- Fixed critical issue where multiple products in a row range were being treated as a single product
- Now correctly identifies and separates different products based on unique descriptions
- Each product is processed independently with its own title, handle, and set of variants
- Handles cases where row ranges (e.g., 14786-14812) contain multiple distinct products
- Ensures that product-specific finishes are correctly applied to each product

## Version 1.5.0 - 2023-06-05 (Finish Prioritization Fix)

### Improvements:
- Fixed issue with finish prioritization for products like Cadiz
- Now properly prioritizes finishes in this order:
  1. Product-specific finishes based on keywords in description (e.g., "Cadiz")
  2. Finishes based on finish count column
  3. Finishes based on finish code (##, x##, etc.)
- This ensures products like "Cadiz Raised Circular Cupboard Knob" will use the 8 Cadiz-specific finishes instead of the 14 standard ## finishes

## Version 1.4.0 - 2023-06-04 (Excel Row Detection Fix)

### Fixes:
- Fixed major issue with Excel row detection where pandas was missing rows
- Added openpyxl integration to accurately detect maximum row number
- Now correctly processes all rows from the Excel file, including rows that pandas couldn't detect
- Verified that data from row 14812 (previously unreachable) is now included in the output

## Version 1.3.1 - 2023-06-04 (Row Range Fix)

### Fixes:
- Improved handling of out-of-range rows when using --rows parameter
- When end row is beyond the last available row, automatically adjusts to use the last available row
- Adds informative warning message showing the adjustment

## Version 1.3.0 - 2023-06-04 (Inventory Policy Update)

### Changes:
- Changed `Variant Inventory Policy` from "continue" to "deny" for all product variants
- No other functionality changes

## Version 1.2.0 - 2023-06-04 (Improved Finish Mapping)

### Improvements:
1. **Smarter SKU-to-Finish Mapping**: 
   - SKUs are now correctly mapped to their specific finishes based on the finish code in MASTER COPY
   - "SCP" and "PB" codes map to their specific finishes
   - "##" code maps to 14 specific finishes (PN, SN, BZ, AB, SB, etc.)
   - "x##" code maps to 8 specific finishes (PCOP, SCOP, BLN, PEW, etc.)

2. **Eliminated Duplicate Size Variants**:
   - Only generates variants for unique sizes, eliminating duplicates
   - Correct number of output rows: unique sizes × applicable finishes

3. **Improved Finish Detection**:
   - Builds a comprehensive mapping of finish codes to full names
   - Correctly identifies which finishes apply to each product row

### Example:
For rows 14770-14774 (1 size with multiple finishes):
- Generated exactly 25 rows (1 size × 25 finishes)
- Mapped each finish to its correct SKU and price:
  - SCP → SKU: 35603, Price: £47.50
  - PB → SKU: 35604, Price: £47.50
  - 14 "##" finishes → SKU: 35605, Price: £47.50
  - 8 "x##" finishes → SKU: 35605X, Price: £59.37

## Version 1.1.0 - 2023-06-04 (Initial Fixes)

### Fixes:
1. **Fixed Excel row handling**: Correctly includes the row above specified range as a header/reference
   - When processing custom rows (e.g., 14770-14774), it now includes row 14769 as reference
   - This ensures all needed information including finish types is captured

2. **Fixed SKU and pricing handling**: 
   - Each product row is now processed with its own unique SKU and price
   - Data is stored by row index instead of by size to preserve all SKU/price variations
   - All variant pricing is correctly preserved from the source data

3. **Fixed variant generation**:
   - Now correctly generates all variants (finishes) for each SKU
   - Generates the correct number of output rows (e.g., 5 SKUs × 25 finishes = 125 rows)

4. **Status field**:
   - Status is now set to "draft" instead of "active"
   - Only set on the first row of each product

### Usage:
```bash
# Process specific rows with custom output file
python3 shopify_feed_generator.py --rows 14770-14774 --output custom_output.xlsx

# Run in test mode with default rows
python3 shopify_feed_generator.py --test

# Process with default output filename (includes timestamp)
python3 shopify_feed_generator.py --rows 14770-14774
```

### Notes:
- A backup of this version has been saved as `shopify_feed_generator_v1.1.0.py`
- The script now handles multiple SKUs and finishes correctly
- Each product's unique SKU/price combination is preserved in the output 