import pandas as pd
import argparse
import glob
import os

def format_output(input_file=None, output_file=None):
    """
    Post-process the generated Shopify feed to fix:
    1. Boolean columns to be "TRUE"/"FALSE" strings
    2. Add price variation for premium finishes
    """
    # If no input file is specified, use the most recent generated file
    if not input_file:
        output_files = glob.glob('shopify_feed_*.xlsx')
        if not output_files:
            print("No output files found")
            return
        input_file = max(output_files, key=os.path.getmtime)
    
    print(f"Processing file: {input_file}")
    
    # Load the file
    df = pd.read_excel(input_file)
    print(f"Loaded {len(df)} rows")
    
    # 2. Add price variation for premium finishes
    # Group by size to get base prices
    size_groups = df.groupby('Option1 Value')
    base_prices = {}
    
    for size, group in size_groups:
        base_prices[size] = group['Variant Price'].min()
    
    # Identify premium finishes
    premium_keywords = ['Polished Nickel', 'Brushed Nickel', 'Antique', 'FFPN', 'FFSN', 'FFAB']
    
    # Create a premium flag column
    df['Is Premium'] = df['Option2 Value'].apply(lambda x: any(keyword in str(x) for keyword in premium_keywords))
    
    # Update prices based on finish type
    for i, row in df.iterrows():
        size = row['Option1 Value']
        finish = str(row['Option2 Value'])
        
        # Check if this is a premium finish
        is_premium = row['Is Premium']
        
        if is_premium:
            # Premium finishes cost 10% more
            df.at[i, 'Variant Price'] = base_prices[size] * 1.1
    
    # Remove the Is Premium column
    df = df.drop(columns=['Is Premium'])
    
    # Save the data to a CSV file first to preserve strings
    temp_csv = 'temp_output.csv'
    df.to_csv(temp_csv, index=False)
    
    # Convert boolean columns to string literals in the CSV
    with open(temp_csv, 'r') as f:
        csv_content = f.read()
    
    # Replace "True" with "TRUE" and "False" with "FALSE"
    csv_content = csv_content.replace(',True,', ',"TRUE",')
    csv_content = csv_content.replace(',False,', ',"FALSE",')
    # Handle edge cases at start or end of line
    csv_content = csv_content.replace(',True\n', ',"TRUE"\n')
    csv_content = csv_content.replace(',False\n', ',"FALSE"\n')
    
    # Write the modified CSV back
    with open(temp_csv, 'w') as f:
        f.write(csv_content)
    
    # Read the CSV back
    df = pd.read_csv(temp_csv)
    
    # Make sure boolean columns are strings
    boolean_columns = ['Published', 'Variant Requires Shipping', 'Variant Taxable', 'Gift Card']
    for col in boolean_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)
            # Replace "nan" with empty string
            df[col] = df[col].replace('nan', '')
            # Make sure TRUE and FALSE are uppercase
            df[col] = df[col].replace('true', 'TRUE').replace('false', 'FALSE')
    
    # Also fix regional inclusion columns
    inclusion_columns = [col for col in df.columns if col.startswith('Included /')]
    for col in inclusion_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)
            # Replace "nan" with empty string
            df[col] = df[col].replace('nan', '')
            # Make sure TRUE is uppercase
            df[col] = df[col].replace('true', 'TRUE')
    
    # Save the updated file
    if not output_file:
        output_file = 'formatted_' + os.path.basename(input_file)
    
    # Save to Excel with string_storage option
    df.to_excel(output_file, index=False)
    print(f"Saved formatted file to {output_file}")
    
    # Clean up temp file
    if os.path.exists(temp_csv):
        os.remove(temp_csv)
    
    # Print summary of changes
    print("\nSummary of changes:")
    
    # Check boolean columns
    for col in boolean_columns:
        if col in df.columns:
            values = df[col].unique()
            print(f"{col} values: {values}")
    
    # Check price variation
    # Add the premium flag back for analysis
    df['Is Premium'] = df['Option2 Value'].apply(lambda x: any(keyword in str(x) for keyword in premium_keywords))
    print("\nPrice variation by premium status:")
    grouped = df.groupby(['Option1 Value', 'Is Premium'])['Variant Price'].mean()
    print(grouped)
    
    # Count premium finishes
    premium_count = df['Is Premium'].sum()
    print(f"\nPremium finishes: {premium_count} out of {len(df)} rows")
    
    return df

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Format Shopify feed output file')
    parser.add_argument('--input', '-i', help='Input Excel file path')
    parser.add_argument('--output', '-o', help='Output Excel file path')
    
    args = parser.parse_args()
    format_output(args.input, args.output) 