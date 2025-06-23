import pandas as pd
import glob
import os

# Find the most recent output file
output_files = glob.glob('shopify_feed_*.xlsx')
if not output_files:
    print("No output files found")
    exit()
    
latest_file = max(output_files, key=os.path.getmtime)
print(f"Checking most recent output file: {latest_file}")

# Load the output file
output_df = pd.read_excel(latest_file)
print(f"Output file has {len(output_df)} rows")

# Group by size and finish to check the distribution
size_finish_counts = output_df.groupby(['Option1 Value', 'Option2 Value']).size().reset_index(name='count')
print("\nDistribution of size and finish combinations:")
print(size_finish_counts)

# Check if all rows have the correct SKUs
size_sku_map = output_df.groupby('Option1 Value')['Variant SKU'].unique()
print("\nSKUs used for each size:")
for size, skus in size_sku_map.items():
    print(f"{size}: {skus}")

# Check if Title, Vendor, Status etc. are only set for the first row
first_row_fields = ['Title', 'Vendor', 'Published', 'Status', 'SEO Title']
for field in first_row_fields:
    non_null_count = output_df[field].notna().sum()
    print(f"\nField '{field}' has {non_null_count} non-null values")
    if non_null_count > 0:
        print(f"Values: {output_df[field].dropna().unique()}")

# Check which finishes are being used
finishes = output_df['Option2 Value'].unique()
print(f"\nFinishes used ({len(finishes)}):")
for i, finish in enumerate(finishes):
    print(f"{i+1}. {finish}")

# Check the data in first few rows
print("\nSample of first few rows:")
sample_columns = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
print(output_df[sample_columns].head(5)) 