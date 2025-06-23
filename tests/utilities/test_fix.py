import pandas as pd
import sys
from shopify_feed_generator import generate_shopify_feed, CONFIG

# Set test rows to 14770-14774
CONFIG["test_start_row"] = 14770
CONFIG["test_end_row"] = 14774

# Generate the feed
print("Generating Shopify feed for rows 14770-14774...")
feed = generate_shopify_feed('MASTER COPY.xlsx', 'test_fixed_output.xlsx', test_mode=True)

# Analyze the output
print("\n=== ANALYSIS OF GENERATED FEED ===")
print(f"Total rows generated: {len(feed)}")

# Count unique handles (products)
unique_handles = feed['Handle'].unique()
print(f"Unique product handles: {len(unique_handles)}")

# Count unique SKUs
unique_skus = feed['Variant SKU'].unique()
print(f"Unique SKUs: {len(unique_skus)}")
print(f"Unique SKUs: {unique_skus}")

# Count rows per SKU
print("\nRows per SKU:")
for sku in unique_skus:
    sku_count = len(feed[feed['Variant SKU'] == sku])
    print(f"SKU {sku}: {sku_count} rows")

# Count size/finish combinations
unique_sizes = feed['Option1 Value'].unique()
unique_finishes = feed['Option2 Value'].unique()
print(f"\nUnique sizes: {len(unique_sizes)}")
print(f"Unique finishes: {len(unique_finishes)}")
print(f"Expected total variants: {len(unique_sizes) * len(unique_finishes)}")

# Print sample rows
print("\nSample rows:")
sample_columns = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
print(feed[sample_columns].head(5))

print("\nTest complete - check test_fixed_output.xlsx for full results") 