import pandas as pd
import glob
import os

# Find the most recent output file
output_files = glob.glob('shopify_feed_*.xlsx')
if not output_files:
    print("No output files found")
    exit()
    
latest_file = max(output_files, key=os.path.getmtime)
print(f"Checking prices in most recent output file: {latest_file}")

# Load the output file
output_df = pd.read_excel(latest_file)
print(f"Output file has {len(output_df)} rows")

# Check prices by size and finish
price_data = output_df[['Option1 Value', 'Option2 Value', 'Variant Price']].sort_values(['Option1 Value', 'Option2 Value'])
print("\nPrices by size and finish:")
print(price_data)

# Analyze price variations
print("\nPrice statistics by size:")
size_price_stats = output_df.groupby('Option1 Value')['Variant Price'].agg(['min', 'max', 'mean', 'std'])
print(size_price_stats)

# Check if premium finishes have higher prices
print("\nChecking if premium finishes have higher prices:")
premium_finishes = ['Polished Nickel', 'Brushed Nickel', 'Antique']
output_df['Is Premium'] = output_df['Option2 Value'].apply(lambda x: any(premium in str(x) for premium in premium_finishes))
print(output_df.groupby(['Option1 Value', 'Is Premium'])['Variant Price'].mean())

# Check if Published column is using TRUE instead of 1
print("\nPublished column values:")
published_values = output_df['Published'].unique()
print(published_values)

# Check other boolean columns
boolean_columns = ['Variant Requires Shipping', 'Variant Taxable', 'Gift Card']
for col in boolean_columns:
    if col in output_df.columns:
        print(f"\n{col} values:")
        print(output_df[col].unique())

# Check regional inclusion columns
inclusion_columns = [col for col in output_df.columns if col.startswith('Included /')]
for col in inclusion_columns:
    print(f"\n{col} values:")
    print(output_df[col].unique()) 