# Changelog for shopify_feed_generator.py

## Version 1.10.0 - 2025-01-15 (Tags and Option Value Enhancements)

### New Features:
- **Tags from Column K**: Automatically populates the "Tags" field in Shopify feed from column K (11th column) in MASTER COPY
  - Tags are extracted from the first row of each product group
  - Only populated in the first row of each product (not every variant)
  - Example: "The Matt Black Suite - Door & Window Hardware"
- **Option1 Value Cleaning**: Automatic cleaning of Size option values for Shopify compatibility
  - Removes commas from Size options and replaces with ` -` (space + hyphen)
  - Example: "60mm Knob, 63mm Base, 70mm P" → "60mm Knob - 63mm Base - 70mm P"
  - Prevents Shopify import issues caused by commas in option values

### Enhanced Error Tracking:
- **Missing Tags Detection**: Products without tags in column K are flagged and not processed
- **Comprehensive Error Reporting**: Missing tags are tracked in `products_not_processed.csv` with reason "Missing tag in column K"
- **Better Quality Control**: Ensures all processed products have proper tag metadata

### Technical Improvements:
- Added `clean_option_value()` function for Size option cleaning
- Added `get_tags_from_column_k()` function for tag extraction from Excel column K
- Applied changes to both test mode and normal processing modes
- Enhanced error tracking with specific tag-related error messages
- Improved data validation to ensure product completeness

### Benefits:
- **Shopify Compatibility**: Cleaned option values prevent import errors
- **Better SEO**: Proper product tags improve searchability and organization
- **Quality Assurance**: Missing tags are caught before feed generation
- **Metadata Completeness**: All products include proper categorization tags

### Example Usage:
```bash
# Test with specific rows that have tags
python3 shopify_feed_generator.py --test --rows "14815-14816"

# Normal processing will now include tags and clean option values
python3 shopify_feed_generator.py --input "MASTER COPY.xlsx" --output "tagged_feed.xlsx"
```

## Version 1.9.0 - 2025-01-15 (Products Without Sizes Support)

### New Features:
- **Products without sizes are now fully supported**: Previously, products without size data were excluded from processing. Now they are included in the Shopify feed.
- **Smart option handling**: Products without sizes use "Finish" as Option1 (instead of Size), with Option2 left empty.
- **Dual processing paths**: The system now intelligently detects whether a product has sizes and processes accordingly:
  - **Products with sizes**: Option1 = "Size", Option2 = "Finish" (existing behavior)
  - **Products without sizes**: Option1 = "Finish", Option2 = empty (new behavior)

### Technical Improvements:
- Modified product filtering logic to include products without size data
- Added `has_size` flag to track whether each row contains size information
- Updated variant creation logic for both test mode and normal mode
- Enhanced logging to clearly indicate when products without sizes are being processed
- Updated error messages to reflect that sizes are no longer required

### Benefits:
- **Increased product coverage**: No products are lost due to missing size data
- **Flexible product structure**: Accommodates different product types (with/without sizes)
- **Backward compatibility**: Products with sizes continue to work exactly as before
- **Better error tracking**: Clear distinction between products without sizes (now valid) vs. products missing essential data (SKU/price)

### Example:
For a product without sizes but with finish code "##":
- Creates 14 variants (one for each finish)
- Option1 Name = "Finish"
- Option1 Value = finish names (e.g., "Polished Nickel (PN)", "Satin Brass (SB)", etc.)
- Option2 Name and Option2 Value remain empty

## Version 1.8.5 - 2025-06-24 (Streamlit Display Fix)

### Fixes:
- Fixed issue where products that couldn't be processed weren't being displayed in the Streamlit UI
- Corrected indentation in app.py to ensure error reports are always visible
- Products with missing data are now properly displayed and can be downloaded as CSV

## Version 1.8.4 - 2025-06-24 (Error Tracking Improvements)

### Improvements:
- Added tracking for products that couldn't be processed due to missing data
- Products with missing size/SKU/price data are now logged to products_not_processed.csv
- Added display and download options for these errors in the Streamlit app
- This helps identify products that were skipped during feed generation

## Version 1.8.3 - 2025-06-24 (X## Finish Code Fix)

### Fixes:
- Fixed case sensitivity issue with X## finish codes
- Premium finishes (X##) now correctly get the SKU from the row with X## finish code
- Ensures all finishes (CP, SCP, PB, ##, and X##) are mapped to their proper SKUs

## Version 1.8.2 - 2025-06-24 (SKU Mapping Fix)

### Fixes:
- Fixed critical issue where all finishes of the same product and size were assigned the same SKU
- Now correctly maps each finish to its specific SKU from the MASTER COPY:
  - Specific finishes (CP, SCP, PB, etc.) get their exact SKU from matching rows
  - Standard finishes (##) get the SKU from the row with ## finish code
  - Premium finishes (x##) get the SKU from the row with x## finish code
- Ensures accurate product codes in the generated Shopify feed

## Version 1.8.1 - 2025-06-24 (Image Alt Text Update)

### Improvements:
- Set "Image Alt Text" to match the "Title" for the first row of each product
- This improves SEO and accessibility for product images

## Version 1.8.0 - 2025-06-24 (Generalization and Compatibility Updates)

### Improvements:
- Removed A&H Brass specific references from app.py to make the tool more generic
- Updated README.md to reflect 25 available finishes (increased from 23)
- Fixed Python 3.13 compatibility issues
- Added security improvements: Explicitly exclude MASTER COPY.xlsx from repository
- Prepared for Streamlit Cloud deployment

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