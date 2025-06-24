# A&H Brass Shopify Product Feed Generator

A Python tool that generates Shopify product feed data from A&H Brass Excel file, with both file upload and manual input options. Now includes automatic tracking of products with unidentified finishes for retrospective fixing.

## Directory Structure

```
.
├── README.md                    # This file - main documentation and usage guide
├── shopify_feed_generator.py    # Main script to generate Shopify product feed
├── app.py                       # Streamlit web application
├── requirements.txt             # Python dependencies
├── run_app.bat                  # Windows batch file to run Streamlit app
├── run_app.sh                   # Mac/Linux shell script to run Streamlit app
├── MASTER COPY.xlsx             # Source data file
├── StreamlitDemo.mp4            # Demo video of the Streamlit app
├── finishes_not_found.csv       # Generated when products have unidentified finishes
├── backup/                      # Backup directory
│   └── versions/                # Previous versions of the script
├── docs/                        # Documentation
│   ├── CHANGELOG.md             # Version history and changes
│   ├── STREAMLIT_GUIDE.md       # Streamlit app usage guide
│   ├── Initial Prompt           # Original project requirements
│   └── Testing & Feedback       # Testing notes and feedback
├── tests/                       # Testing directory
│   ├── utilities/               # Test utilities and scripts
│   └── output/                  # Test output files
└── __pycache__/                 # Python cache (generated automatically)
```

## Usage

### Command Line Interface

```bash
# Basic usage (processes all new products)
python3 shopify_feed_generator.py

# Process specific rows with a custom output file
python3 shopify_feed_generator.py --rows 14786-14812 --output custom_output.xlsx

# Run in test mode with default rows
python3 shopify_feed_generator.py --test

# Show version information
python3 shopify_feed_generator.py --version
```

### Web Interface (Streamlit App)

For a more user-friendly experience, you can use the Streamlit web app:

```bash
# Install dependencies
pip install -r requirements.txt

# Run the Streamlit app
streamlit run app.py
```

The web interface provides two methods for generating Shopify feeds:

#### Method 1: File Upload
- Upload your MASTER COPY.xlsx file
- Select row ranges visually
- Preview products before generating
- Generate and download the Shopify feed with a simple click

#### Method 2: Manual Input
- Enter product data manually using the same structure as the Excel file
- Input rows with: Description, Size, SKU, Price, Finish Code, Finish Count
- Uses the exact same processing logic as the file upload method
- Supports finish codes: `##` (14 finishes), `x##` (8 finishes), or specific codes
- Perfect for users who would normally enter data into the Excel file

## Features

- **Two input methods**: File upload or manual data entry
- Generates Shopify product feed from MASTER COPY Excel file
- **Manual product creation**: Enter product data through web forms
- **Finish selection**: Choose from 25 available finishes including:
  - Factory Finished Satin Brass (FFSB)
  - Factory Finished Polished Nickel (FFPN)
  - Factory Finished Matt Bronze (FFMB)
  - And 20 more finish options
- **Finishes tracking and reporting**: Automatically identifies products with unidentified finishes
- Correctly handles multiple products with different variants
- Properly prioritizes finishes based on product names (e.g., Cadiz)
- Supports custom row selection for targeted processing
- Handles finish codes (##, x##) for product variants
- Produces properly formatted Excel output ready for Shopify import
- **Dual-tab interface**: File upload and manual input don't interfere with each other

## Finishes Tracking

The system automatically tracks products where finishes couldn't be identified and had to default to using all 25 finishes. This helps with retrospective fixing and quality control.

### Features:
- **Automatic Detection**: Identifies products with unknown finish codes
- **CSV Export**: Creates `finishes_not_found.csv` with detailed information when issues are found
- **Streamlit Integration**: Shows warnings and downloadable reports in the web app
- **Console Feedback**: Clear success/warning messages in command line

### Output Information:
When products with unidentified finishes are found, the system exports a CSV containing:
- Product Description
- Row Index
- Size & SKU
- Finish Code that couldn't be identified
- Reason for defaulting
- What it defaulted to (e.g., "25 finishes from column 25")

### Example Output:
```bash
✅ All products had identifiable finishes
# OR
⚠️ Found 3 products with unidentified finishes
📄 Details exported to finishes_not_found.csv
```

## Manual Input Structure

When using the manual input method, you enter data row by row as it would appear in the Excel file:

### Row Fields
- **Description**: Product name (same for all sizes of one product)
- **Size**: Size description (e.g., "32mm x 32mm p")
- **SKU/Code**: Product SKU (e.g., "35607/1")
- **Price**: Retail price in GBP
- **Finish Code**: Finish specification:
  - `##` = Applies to 14 standard finishes
  - `x##` = Applies to 8 premium finishes
  - Specific codes like `FFSB`, `FFPN` = Applies to that specific finish only
  - Empty = Uses all available finishes
- **Finish Count**: Optional number for product-specific finish detection

### Example
For a Cadiz knob with 2 sizes, you would enter 2 rows:
1. Cadiz Raised Circular Cupboard Knob | 32mm x 32mm p | 35607/1 | 18.55 | ## | 8
2. Cadiz Raised Circular Cupboard Knob | 38mm x 32mm p | 35607/2 | 20.58 | ## | 8

This generates 16 variants (2 sizes × 8 finishes) using the same logic as the file upload method.

## Latest Version

Version 1.7.1 includes:
- **Finishes tracking and reporting** for products with unidentified finishes
- **Enhanced column F logic** for better finish selection before defaulting
- **CSV export functionality** for retrospective fixing (`finishes_not_found.csv`)
- **Improved Streamlit UI** with finish status indicators and downloadable reports
- **New manual input functionality** with comprehensive form interface
- **Dual-tab design** - file upload and manual input in separate tabs
- **23 predefined finishes** with proper naming conventions
- **Dynamic size management** - add/remove sizes as needed
- **Form validation** to ensure required fields are completed
- Multiple product support (correctly separates different products in a row range)
- Finish prioritization based on product name and finish count
- Accurate Excel row detection using openpyxl

For complete version history, see the [CHANGELOG.md](docs/CHANGELOG.md) file. 