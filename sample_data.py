"""
Sample data generator for A&H Brass Shopify Feed Generator
This creates sample Excel data when MASTER COPY.xlsx is not available
"""
import pandas as pd
import os

def create_sample_excel():
    """Create a sample Excel file with minimal data for demo purposes"""
    
    # Sample MASTER COPY data
    master_copy_data = {
        'description': [
            'Sample Lever Handle',
            'Sample Lever Handle', 
            'Sample Cupboard Knob',
            'Sample Cupboard Knob'
        ],
        'size': [
            '150mm x 50mm',
            '200mm x 50mm',
            '25mm',
            '32mm'
        ],
        'code': [
            'SMP001/1',
            'SMP001/2', 
            'SMP002/1',
            'SMP002/2'
        ],
        'rrp': [25.50, 28.50, 15.50, 18.50],
        'finish': ['##', '##', 'x##', 'x##'],
        'finish count': [14, 14, 8, 8]
    }
    
    # Sample finishes data
    finishes_data = {
        '14': [
            'Factory Finished Polished Nickel (PN)',
            'Factory Finished Satin Nickel (SN)', 
            'Factory Finished Bronze (BZ)',
            'Factory Finished Antique Brass (AB)',
            'Factory Finished Satin Brass (SB)',
            'Factory Finished Dark Bronze (DB)',
            'Factory Finished Black Antique Brass (BAB)',
            'Factory Finished Bronze Waxed (BZW)',
            'Factory Finished Black Antique Brass Waxed (BABW)',
            'Factory Finished Antique Brass Waxed (ABW)',
            'Factory Finished Dark Bronze Waxed (DBW)',
            'Factory Finished Nickel Bronze Waxed (NBW)',
            'Factory Finished Satin Brass Waxed (SBW)',
            'Factory Finished Polished Brass Unlacquered (PBUL)'
        ],
        '8': [
            'Factory Finished Polished Copper (PCOP)',
            'Factory Finished Satin Copper (SCOP)',
            'Factory Finished Black Nickel (BLN)',
            'Factory Finished Pewter (PEW)',
            'Factory Finished Matt Black (MBL)',
            'Factory Finished Antique Silver (ASV)',
            'Factory Finished Rose Gold Plated (RGP)',
            'Factory Finished Aged Copper (ACOP)'
        ],
        '25': [
            'Factory Finished Polished Nickel (PN)',
            'Factory Finished Satin Nickel (SN)', 
            'Factory Finished Bronze (BZ)',
            'Factory Finished Antique Brass (AB)',
            'Factory Finished Satin Brass (SB)',
            'Factory Finished Dark Bronze (DB)',
            'Factory Finished Black Antique Brass (BAB)',
            'Factory Finished Bronze Waxed (BZW)',
            'Factory Finished Black Antique Brass Waxed (BABW)',
            'Factory Finished Antique Brass Waxed (ABW)',
            'Factory Finished Dark Bronze Waxed (DBW)',
            'Factory Finished Nickel Bronze Waxed (NBW)',
            'Factory Finished Satin Brass Waxed (SBW)',
            'Factory Finished Polished Brass Unlacquered (PBUL)',
            'Factory Finished Polished Copper (PCOP)',
            'Factory Finished Satin Copper (SCOP)',
            'Factory Finished Black Nickel (BLN)',
            'Factory Finished Pewter (PEW)',
            'Factory Finished Matt Black (MBL)',
            'Factory Finished Antique Silver (ASV)',
            'Factory Finished Rose Gold Plated (RGP)',
            'Factory Finished Aged Copper (ACOP)',
            'Factory Finished Matt Bronze (FFMB)',
            'Factory Finished Matt Brass (FFMBL)',
            'Factory Finished Chrome (CHR)'
        ]
    }
    
    # Sample template data
    sample_data = {
        'Handle': ['sample-product', 'sample-product'],
        'Title': ['Sample Product', None],
        'Option1 Name': ['Size', None],
        'Option1 Value': ['25mm', '32mm'],
        'Option2 Name': ['Finish', None],
        'Option2 Value': ['Factory Finished Polished Nickel (PN)', 'Factory Finished Satin Nickel (SN)'],
        'Variant SKU': ['SMP001/1', 'SMP001/2'],
        'Variant Price': [25.50, 28.50]
    }
    
    # Create DataFrames
    master_copy_df = pd.DataFrame(master_copy_data)
    finishes_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in finishes_data.items()]))
    sample_df = pd.DataFrame(sample_data)
    
    # Save to Excel file
    filename = 'SAMPLE_MASTER_COPY.xlsx'
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        master_copy_df.to_excel(writer, sheet_name='MASTER COPY', index=False)
        finishes_df.to_excel(writer, sheet_name='Finishes', index=False)
        sample_df.to_excel(writer, sheet_name='Sample', index=False)
    
    print(f"Sample Excel file created: {filename}")
    return filename

if __name__ == "__main__":
    create_sample_excel() 