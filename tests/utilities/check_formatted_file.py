import pandas as pd
import glob
import os

# Load the formatted file directly
formatted_file = 'formatted_shopify_feed_20250604_120309.xlsx'
print(f"Checking formatted file: {formatted_file}")

# Load the file
output_df = pd.read_excel(formatted_file)
print(f"Output file has {len(output_df)} rows")

# Check boolean columns
boolean_columns = ['Published', 'Variant Requires Shipping', 'Variant Taxable', 'Gift Card']
for col in boolean_columns:
    if col in output_df.columns:
        values = output_df[col].unique()
        print(f"\n{col} values:")
        print(values)
        # Count non-null values
        non_null_count = output_df[col].notna().sum()
        print(f"Non-null count: {non_null_count}")
        # Count TRUE values
        true_count = (output_df[col] == "TRUE").sum()
        print(f"TRUE count: {true_count}")

# Check regional inclusion columns
inclusion_columns = [col for col in output_df.columns if col.startswith('Included /')]
for col in inclusion_columns:
    values = output_df[col].unique()
    print(f"\n{col} values:")
    print(values)
    # Count TRUE values
    true_count = (output_df[col] == "TRUE").sum()
    print(f"TRUE count: {true_count}")

# Check price variation by finish
# Add the premium flag back for analysis
premium_keywords = ['Polished Nickel', 'Brushed Nickel', 'Antique', 'FFPN', 'FFSN', 'FFAB']
output_df['Is Premium'] = output_df['Option2 Value'].apply(lambda x: any(keyword in str(x) for keyword in premium_keywords))

# Group by size and premium status
print("\nPrice variation by premium status:")
grouped = output_df.groupby(['Option1 Value', 'Is Premium'])['Variant Price'].agg(['min', 'max', 'mean'])
print(grouped)

# Compare prices by finish
print("\nPrices by size and finish:")
price_data = output_df[['Option1 Value', 'Option2 Value', 'Variant Price', 'Is Premium']].sort_values(['Option1 Value', 'Option2 Value'])
print(price_data) 