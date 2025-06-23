# Streamlit App Guide

The A&H Brass Shopify Feed Generator app provides a user-friendly web interface for generating Shopify product feeds from your MASTER COPY Excel file. This guide walks you through using the app effectively.

## Getting Started

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Launch the app:
   ```bash
   streamlit run app.py
   ```

3. The app will open in your default web browser (typically at http://localhost:8501)

## Using the App

### Step 1: Upload Your Excel File

- Click the "Upload your MASTER COPY.xlsx file" button
- Select your Excel file containing the product data
- The app will verify that the file contains the required sheets (MASTER COPY, Sample, Finishes)

### Step 2: Select Rows to Process

- Use the number input fields to specify the start and end rows
- Default values are set to rows 14786-14812 (for example)
- The app will show a preview of products found in this range

### Step 3: Preview Products

- The app will display:
  - A bar chart showing the distribution of products
  - A detailed list of products with their row numbers
  - This helps you verify you've selected the correct range

### Step 4: Generate the Feed

- Click the "Generate Shopify Feed" button
- The app will process the data (this may take a moment)
- You'll see statistics about the generated feed (products, variants, etc.)

### Step 5: Download the Result

- A preview of the generated feed will be displayed
- Click the download link to save the Excel file to your computer
- The file will be named with a timestamp (e.g., shopify_feed_20230605_123045.xlsx)

## Troubleshooting

- **Missing Sheets**: If your Excel file is missing required sheets (MASTER COPY, Sample, Finishes), you'll see an error message.
- **No Products Found**: If no products are detected in your selected row range, try adjusting the range.
- **Generation Errors**: If you encounter errors during generation, check that your Excel file matches the expected format.

## Advanced Usage

- The app automatically detects multiple products in your selected row range
- It prioritizes finishes based on product name keywords (like "Cadiz")
- All variants are correctly generated with appropriate finish codes and SKUs

---

For technical details about how the generator works, refer to the [README.md](../README.md) file. 