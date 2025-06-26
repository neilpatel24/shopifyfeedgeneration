import streamlit as st
import pandas as pd
import io
import os
import sys
import tempfile
import base64
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import openpyxl  # Import openpyxl for accurate row detection

# Import the shopify_feed_generator module
from shopify_feed_generator import generate_shopify_feed, __version__, CONFIG

# Set a nice color palette for charts
plt.style.use('ggplot')
COLORS = ["#1E88E5", "#FFC107", "#26A69A", "#D81B60", "#8E24AA", "#E53935", "#43A047"]

st.set_page_config(
    page_title="Shopify Feed Generator",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS
st.markdown("""
<style>
    .main .block-container {
        padding-top: 2rem;
    }
    h1, h2, h3 {
        color: #1E3A8A;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
    }
    .stDownloadButton>button {
        background-color: #15803D;
        color: white;
    }
    .highlight {
        background-color: #F0F9FF;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1E3A8A;
    }
    .manual-input-form {
        background-color: #F8FAFC;
        padding: 1.5rem;
        border-radius: 0.5rem;
        border: 1px solid #E2E8F0;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Define available finishes
AVAILABLE_FINISHES = {
    'FFSB': 'Factory Finished Satin Brass (FFSB)',
    'FFPN': 'Factory Finished Polished Nickel (FFPN)',
    'FFMB': 'Factory Finished Matt Bronze (FFMB)',
    'FFAB': 'Factory Finished Antique Brass (FFAB)',
    'FFSN': 'Factory Finished Satin Nickel (FFSN)',
    'FFDB': 'Factory Finished Dark Bronze (FFDB)',
    'FFBAB': 'Factory Finished Burnished Antique Brass (FFBAB)',
    'FFBZW': 'Factory Finished Bronze Wax (FFBZW)',
    'FFBABW': 'Factory Finished Burnished Antique Brass Wax (FFBABW)',
    'FFABW': 'Factory Finished Antique Brass Wax (FFABW)',
    'FFDBW': 'Factory Finished Dark Bronze Wax (FFDBW)',
    'FFNBW': 'Factory Finished Natural Brass Wax (FFNBW)',
    'FFSBW': 'Factory Finished Satin Brass Wax (FFSBW)',
    'FFPBUL': 'Factory Finished Polished Brass Unlacquered (FFPBUL)',
    'FFPCOP': 'Factory Finished Polished Copper (FFPCOP)',
    'FFSCOP': 'Factory Finished Satin Copper (FFSCOP)',
    'FFBLN': 'Factory Finished Black Nickel (FFBLN)',
    'FFPEW': 'Factory Finished Pewter (FFPEW)',
    'FFMBL': 'Factory Finished Matt Black (FFMBL)',
    'FFASV': 'Factory Finished Antique Silver (FFASV)',
    'FFRGP': 'Factory Finished Rose Gold Plated (FFRGP)',
    'FFACOP': 'Factory Finished Antique Copper (FFACOP)'
}

def get_excel_download_link(df, filename):
    """Generate a download link for an Excel file from a DataFrame"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    b64 = base64.b64encode(output.getvalue()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def get_excel_preview(excel_file):
    """Get a preview of sheets in the Excel file"""
    try:
        xls = pd.ExcelFile(excel_file)
        sheets = xls.sheet_names
        
        st.write("### Excel File Structure")
        st.write(f"Found {len(sheets)} sheets: {', '.join(sheets)}")
        
        # Check for required sheets
        required_sheets = ['MASTER COPY', 'Sample', 'Finishes']
        missing_sheets = [sheet for sheet in required_sheets if sheet not in sheets]
        
        if missing_sheets:
            st.error(f"⚠️ Missing required sheets: {', '.join(missing_sheets)}")
            return None
        
        # Use openpyxl to get accurate row count
        wb = openpyxl.load_workbook(excel_file, read_only=True)
        sheet = wb['MASTER COPY']
        max_row = sheet.max_row
        wb.close()
        
        # Read the MASTER COPY sheet with pandas
        df = pd.read_excel(excel_file, sheet_name='MASTER COPY')
        
        # Add a note about the row count
        st.write(f"📊 Excel file contains {max_row} rows in MASTER COPY sheet")
        
        # Check if there are any rows with descriptions
        has_descriptions = not df['description'].dropna().empty if 'description' in df.columns else False
        
        if not has_descriptions:
            st.error("⚠️ No product descriptions found in the MASTER COPY sheet")
            return None
        
        return {
            'df': df,
            'max_row': max_row,  # Use openpyxl's accurate max_row
            'sheets': sheets,
            'has_descriptions': has_descriptions
        }
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def analyze_products(df, start_row, end_row):
    """Analyze products in the specified row range"""
    # Make a copy of the dataframe and adjust indexing to match Excel row numbers
    df = df.copy()
    
    # Filter the dataframe to get rows we're interested in
    # Pandas is 0-indexed, but Excel is 1-indexed
    # We need to convert between the two and apply a +1 offset to match the generator's row numbering
    
    # Convert dataframe index to match Excel row numbers with +1 offset
    excel_rows = {}
    for i, (idx, row) in enumerate(df.iterrows()):
        # Excel row number is pandas index + 1, and we need to add 1 more to match the generator
        excel_row_num = idx + 2  # Add 2 instead of 1 to fix the off-by-one issue
        excel_rows[excel_row_num] = row
    
    # Now group by description for rows in our range
    product_groups = {}
    current_description = None
    
    # Process each row in our selected range
    for row_num in range(start_row, end_row + 1):
        # Skip if row doesn't exist in our dataframe
        if row_num not in excel_rows:
            continue
            
        row = excel_rows[row_num]
        description = row.get('description')
        
        # If this row has a description
        if not pd.isna(description) and isinstance(description, str) and len(description) > 0:
            current_description = description
            if description not in product_groups:
                product_groups[description] = []
            product_groups[description].append(row_num)
        
        # Row without description but with data and following a known product
        elif current_description and (
            not pd.isna(row.get('size')) or 
            not pd.isna(row.get('code')) or 
            not pd.isna(row.get('rrp'))
        ):
            product_groups[current_description].append(row_num)
    
    return product_groups

def plot_product_distribution(products):
    """Create a visualization of product distribution"""
    product_counts = {k: len(v) for k, v in products.items()}
    
    # If no products or all empty, return None
    if not product_counts or all(count == 0 for count in product_counts.values()):
        return None
        
    # Sort products by count for better visualization
    sorted_products = dict(sorted(product_counts.items(), key=lambda item: item[1], reverse=True))
    
    # Create a more compact and visually appealing chart
    fig, ax = plt.subplots(figsize=(6, min(5, 1 + len(sorted_products) * 0.4)))  # Smaller, adaptive height
    
    # Create a horizontal bar chart
    if len(sorted_products) > 0:
        y_pos = range(len(sorted_products))
        product_names = list(sorted_products.keys())
        counts = list(sorted_products.values())
        
        # Truncate long product names
        product_names = [p[:25] + '...' if len(p) > 25 else p for p in product_names]
        
        # Plot horizontal bars with a color palette
        bars = ax.barh(y_pos, counts, align='center', 
                       color=COLORS[:len(sorted_products)] if len(sorted_products) <= len(COLORS) 
                       else plt.cm.tab10(range(len(sorted_products))), 
                       alpha=0.8,
                       height=0.6)  # Thinner bars
        
        # Customize the plot
        ax.set_yticks(y_pos)
        ax.set_yticklabels(product_names, fontsize=9)
        ax.invert_yaxis()  # Labels read top-to-bottom
        ax.set_xlabel('Number of Rows', fontsize=9)
        ax.set_title('Products Distribution', fontsize=11, fontweight='bold')
        
        # Remove spines for cleaner look
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        # Add count labels to the right of each bar
        for i, v in enumerate(counts):
            ax.text(v + 0.1, i, str(v), va='center', fontsize=9, fontweight='bold')
        
        plt.tight_layout()
        return fig
    return None

def get_excel_file_path():
    """Get the path to the Excel file, using sample file as fallback for deployment"""
    if os.path.exists('MASTER COPY.xlsx'):
        return 'MASTER COPY.xlsx'
    elif os.path.exists('SAMPLE_MASTER_COPY.xlsx'):
        return 'SAMPLE_MASTER_COPY.xlsx'
    else:
        return None

def create_manual_shopify_feed(manual_rows_data):
    """Create a Shopify feed from manually entered product data using the same logic as file upload"""
    
    # Create a DataFrame from the manual input that mimics the Excel structure
    df_data = []
    for row in manual_rows_data:
        df_data.append({
            'description': row['description'],
            'size': row['size'], 
            'code': row['sku'],
            'rrp': row['price'],
            'finish': row['finish_code'],
            'finish count': row.get('finish_count', None)
        })
    
    manual_df = pd.DataFrame(df_data)
    
    # Create temporary files to mimic the Excel structure needed by the generator
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
        # Create a workbook with the required sheets
        with pd.ExcelWriter(tmp_file.name, engine='openpyxl') as writer:
            # Create MASTER COPY sheet with our manual data
            manual_df.to_excel(writer, sheet_name='MASTER COPY', index=False)
            
            # Create Sample sheet (copy from existing file if available)
            try:
                excel_file = get_excel_file_path()
                if excel_file:
                    sample_df = pd.read_excel(excel_file, sheet_name='Sample')
                    sample_df.to_excel(writer, sheet_name='Sample', index=False)
                else:
                    # Create a minimal sample sheet
                    sample_columns = ['Handle', 'Title', 'Option1 Name', 'Option1 Value', 'Option2 Name', 'Option2 Value', 'Variant SKU', 'Variant Price']
                    empty_sample = pd.DataFrame(columns=sample_columns)
                    empty_sample.to_excel(writer, sheet_name='Sample', index=False)
            except:
                # Create a minimal sample sheet
                sample_columns = ['Handle', 'Title', 'Option1 Name', 'Option1 Value', 'Option2 Name', 'Option2 Value', 'Variant SKU', 'Variant Price']
                empty_sample = pd.DataFrame(columns=sample_columns)
                empty_sample.to_excel(writer, sheet_name='Sample', index=False)
            
            # Create Finishes sheet (copy from existing file if available)
            try:
                excel_file = get_excel_file_path()
                if excel_file:
                    finishes_df = pd.read_excel(excel_file, sheet_name='Finishes')
                    finishes_df.to_excel(writer, sheet_name='Finishes', index=False)
                else:
                    # Create a minimal finishes sheet
                    finishes_data = {
                        '8': list(AVAILABLE_FINISHES.values())[:8],
                        '25': list(AVAILABLE_FINISHES.values())
                    }
                    finishes_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in finishes_data.items()]))
                    finishes_df.to_excel(writer, sheet_name='Finishes', index=False)
            except:
                # Create a minimal finishes sheet
                finishes_data = {
                    '8': list(AVAILABLE_FINISHES.values())[:8],
                    '25': list(AVAILABLE_FINISHES.values())
                }
                finishes_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in finishes_data.items()]))
                finishes_df.to_excel(writer, sheet_name='Finishes', index=False)
        
        # Now use the existing generator logic with test mode
        CONFIG["test_start_row"] = 1  # Start from first row of our manual data
        CONFIG["test_end_row"] = len(manual_df)  # End at last row of our manual data
        
        try:
            feed_df, finishes_not_found, products_not_processed = generate_shopify_feed(tmp_file.name, test_mode=True)
            return feed_df, finishes_not_found, products_not_processed
        finally:
            # Clean up the temporary file
            os.unlink(tmp_file.name)

def main():
    st.title("Shopify Feed Generator")
    st.write(f"Version {__version__}")
    
    # Create tabs
    tab1, tab2 = st.tabs(["📁 File Upload", "✏️ Manual Input"])
    
    with tab1:
        st.header("Upload Excel File Method")
        
        with st.sidebar:
            st.header("Instructions")
            st.markdown("""
            1. Upload your **MASTER COPY.xlsx** file
            2. Specify the row range to process
            3. Preview the products detected
            4. Generate the Shopify feed
            5. Download the result
            """)
            
            st.divider()
            
            st.header("About")
            st.markdown("""
            This tool converts product data from Excel to Shopify-compatible format.
            
            Features:
            - Multi-product detection
            - Finish prioritization
            - Variant generation
            """)
        
        # File upload
        uploaded_file = st.file_uploader("Upload your MASTER COPY.xlsx file", type=["xlsx"])
        
        if uploaded_file is not None:
            # Save the uploaded file to a temporary location
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # Get Excel preview
            preview_data = get_excel_preview(tmp_file_path)
            
            if preview_data:
                df = preview_data['df']
                max_row = preview_data['max_row']
                
                # Row selection
                st.write("### Select Rows to Process")
                col1, col2 = st.columns(2)
                
                with col1:
                    start_row = st.number_input("Start Row", min_value=1, max_value=max_row, value=1)
                
                with col2:
                    end_row = st.number_input("End Row", min_value=start_row, max_value=max_row, value=min(100000, max_row))
                
                # Analyze products in the selected range
                products = analyze_products(df, start_row, end_row)
                
                if products:
                    st.write(f"### Found {len(products)} Products in Rows {start_row}-{end_row}")
                    
                    # Create two columns for better space usage
                    col1, col2 = st.columns([2, 3])
                    
                    with col1:
                        # Plot product distribution
                        fig = plot_product_distribution(products)
                        if fig:
                            st.pyplot(fig)
                    
                    with col2:
                        # Show details for each product
                        with st.expander("Product Details", expanded=True):
                            for product_name, rows in products.items():
                                st.markdown(f"**{product_name}**")
                                st.write(f"Rows: {', '.join(map(str, rows[:5]))}{' ...' if len(rows) > 5 else ''}")
                                st.write(f"Total rows: {len(rows)}")
                                st.divider()
                    
                    # Generate button
                    if st.button("Generate Shopify Feed", key="file_upload_generate"):
                        with st.spinner("Generating Shopify feed..."):
                            # Set the configuration for the generator
                            CONFIG["test_start_row"] = start_row
                            CONFIG["test_end_row"] = end_row
                            
                            # Create output filename
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            output_file = f"shopify_feed_{timestamp}.xlsx"
                            
                            try:
                                # Run the generator
                                feed_df, finishes_not_found, products_not_processed = generate_shopify_feed(tmp_file_path, output_file, test_mode=True)
                                
                                if not feed_df.empty:
                                    st.success(f"✅ Successfully generated Shopify feed with {len(feed_df)} rows!")
                                    
                                    # Show finishes not found warning if any
                                    if finishes_not_found:
                                        st.warning(f"⚠️ Found {len(finishes_not_found)} products with unidentified finishes")
                                        
                                        with st.expander("🔍 View Products with Unidentified Finishes", expanded=False):
                                            finishes_df = pd.DataFrame(finishes_not_found)
                                            st.dataframe(finishes_df, use_container_width=True)
                                            
                                            # Offer download of the CSV
                                            csv = finishes_df.to_csv(index=False).encode('utf-8')
                                            st.download_button(
                                                label="Download Finishes Not Found Report",
                                                data=csv,
                                                file_name="finishes_not_found.csv",
                                                mime="text/csv"
                                            )
                                        
                                        # Show products not processed warning if any
                                        if products_not_processed:
                                            st.warning(f"⚠️ Found {len(products_not_processed)} products that couldn't be processed")
                                            
                                            with st.expander("🔍 View Products That Couldn't Be Processed", expanded=False):
                                                not_processed_df = pd.DataFrame(products_not_processed)
                                                st.dataframe(not_processed_df, use_container_width=True)
                                                
                                                # Offer download of the CSV
                                                csv = not_processed_df.to_csv(index=False).encode('utf-8')
                                                st.download_button(
                                                    label="Download Products Not Processed Report",
                                                    data=csv,
                                                    file_name="products_not_processed.csv",
                                                    mime="text/csv"
                                                )
                                    else:
                                        st.info("✅ All products had identifiable finishes")
                                    
                                    # Count unique products and variants
                                    unique_handles = feed_df['Handle'].unique()
                                    
                                    with st.container():
                                        st.markdown('<div class="highlight">', unsafe_allow_html=True)
                                        st.write(f"📊 Generated {len(unique_handles)} unique products with {len(feed_df)} total variants")
                                        
                                        # Create metrics for products and variants
                                        col1, col2, col3 = st.columns(3)
                                        with col1:
                                            st.metric("Products", len(unique_handles))
                                        with col2:
                                            st.metric("Variants", len(feed_df))
                                        with col3:
                                            st.metric("Avg. Variants per Product", round(len(feed_df) / len(unique_handles), 1))
                                        
                                        st.markdown('</div>', unsafe_allow_html=True)
                                    
                                    # Preview of the generated feed
                                    st.write("### Preview of Shopify Feed")
                                    preview_cols = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
                                    st.dataframe(feed_df[preview_cols].head(20), use_container_width=True)
                                    
                                    # Download link
                                    st.markdown(get_excel_download_link(feed_df, output_file), unsafe_allow_html=True)
                                else:
                                    st.error("❌ Failed to generate Shopify feed. The output is empty.")
                            except Exception as e:
                                st.error(f"❌ Error generating Shopify feed: {e}")
                else:
                    st.warning("⚠️ No products found in the selected row range. Please select a different range.")
            
            # Clean up the temporary file
            os.unlink(tmp_file_path)
    
    with tab2:
        st.header("Manual Product Input Method")
        st.write("Enter product data manually using the same structure as the Excel file")
        
        st.info("""
        💡 **How this works**: Enter data as if you were filling out rows in the Excel file. Each row represents a size/finish combination.
        The system will use the same logic as the file upload method to group products and generate variants.
        """)
        
        # Initialize session state for manual rows
        if 'manual_rows' not in st.session_state:
            st.session_state.manual_rows = [
                {
                    'description': '',
                    'size': '',
                    'sku': '',
                    'price': 0.0,
                    'finish_code': '',
                    'finish_count': None
                }
            ]
        
        # Form for manual input
        with st.form("manual_input_form"):
            st.subheader("📝 Product Data Rows")
            st.write("Enter each row as it would appear in the Excel file:")
            
            # Display each row
            for i, row in enumerate(st.session_state.manual_rows):
                st.markdown(f"**Row {i+1}**")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    row['description'] = st.text_input(
                        "Product Description", 
                        value=row['description'], 
                        key=f"desc_{i}",
                        placeholder="e.g., Cadiz Raised Circular Cupboard Knob",
                        help="Product name (same for all sizes of one product)"
                    )
                    row['size'] = st.text_input(
                        "Size", 
                        value=row['size'], 
                        key=f"size_{i}",
                        placeholder="e.g., 32mm x 32mm p",
                        help="Size description"
                    )
                
                with col2:
                    row['sku'] = st.text_input(
                        "SKU/Code", 
                        value=row['sku'], 
                        key=f"sku_{i}",
                        placeholder="e.g., 35607/1",
                        help="Product SKU code"
                    )
                    row['price'] = st.number_input(
                        "Price (£)", 
                        value=row['price'], 
                        key=f"price_{i}",
                        min_value=0.0, 
                        step=0.01,
                        help="Retail price"
                    )
                
                with col3:
                    finish_options = ['', '##', 'x##'] + list(AVAILABLE_FINISHES.keys())
                    current_finish = row['finish_code'] if row['finish_code'] in finish_options else ''
                    row['finish_code'] = st.selectbox(
                        "Finish Code", 
                        options=finish_options,
                        index=finish_options.index(current_finish),
                        key=f"finish_{i}",
                        help="## = 14 standard finishes, x## = 8 premium finishes, or specific finish code"
                    )
                    row['finish_count'] = st.number_input(
                        "Finish Count", 
                        value=row['finish_count'] if row['finish_count'] is not None else 0, 
                        key=f"finish_count_{i}",
                        min_value=0,
                        help="Number of finishes (optional, for product-specific finish detection)"
                    )
                    if row['finish_count'] == 0:
                        row['finish_count'] = None
                
                st.divider()
            
            # Form buttons
            col1, col2, col3 = st.columns([1, 1, 3])
            
            with col1:
                add_row = st.form_submit_button("➕ Add Row")
            
            with col2:
                remove_row = st.form_submit_button("➖ Remove Last Row")
            
            # Main submit button
            submitted = st.form_submit_button("🚀 Generate Shopify Feed", type="primary")
        
        # Handle form actions
        if add_row:
            st.session_state.manual_rows.append({
                'description': '',
                'size': '',
                'sku': '',
                'price': 0.0,
                'finish_code': '',
                'finish_count': None
            })
            st.rerun()
        
        if remove_row and len(st.session_state.manual_rows) > 1:
            st.session_state.manual_rows.pop()
            st.rerun()
        
        if submitted:
            # Validate the data
            errors = []
            valid_rows = []
            
            for i, row in enumerate(st.session_state.manual_rows):
                if not row['description']:
                    if any([row['size'], row['sku'], row['price']]):  # If other fields filled but no description
                        errors.append(f"Row {i+1}: Description is required when other fields are filled")
                elif not row['size']:
                    errors.append(f"Row {i+1}: Size is required")
                elif not row['sku']:
                    errors.append(f"Row {i+1}: SKU is required")
                elif row['price'] <= 0:
                    errors.append(f"Row {i+1}: Price must be greater than 0")
                else:
                    valid_rows.append(row)
            
            if not valid_rows:
                errors.append("At least one complete row is required")
            
            if errors:
                for error in errors:
                    st.error(f"❌ {error}")
            else:
                try:
                    with st.spinner("Generating Shopify feed using the same logic as file upload..."):
                        feed_df, finishes_not_found, products_not_processed = create_manual_shopify_feed(valid_rows)
                        
                        if not feed_df.empty:
                            st.success(f"✅ Successfully generated Shopify feed with {len(feed_df)} variants!")
                            
                            # Show finishes not found warning if any
                            if finishes_not_found:
                                st.warning(f"⚠️ Found {len(finishes_not_found)} products with unidentified finishes")
                                
                                with st.expander("🔍 View Products with Unidentified Finishes", expanded=False):
                                    finishes_df = pd.DataFrame(finishes_not_found)
                                    st.dataframe(finishes_df, use_container_width=True)
                                    
                                    # Offer download of the CSV
                                    csv = finishes_df.to_csv(index=False).encode('utf-8')
                                    st.download_button(
                                        label="Download Finishes Not Found Report",
                                        data=csv,
                                        file_name="finishes_not_found.csv",
                                        mime="text/csv"
                                    )
                                
                                # Show products not processed warning if any
                                if products_not_processed:
                                    st.warning(f"⚠️ Found {len(products_not_processed)} products that couldn't be processed")
                                    
                                    with st.expander("🔍 View Products That Couldn't Be Processed", expanded=False):
                                        not_processed_df = pd.DataFrame(products_not_processed)
                                        st.dataframe(not_processed_df, use_container_width=True)
                                        
                                        # Offer download of the CSV
                                        csv = not_processed_df.to_csv(index=False).encode('utf-8')
                                        st.download_button(
                                            label="Download Products Not Processed Report",
                                            data=csv,
                                            file_name="products_not_processed.csv",
                                            mime="text/csv"
                                        )
                            else:
                                st.info("✅ All products had identifiable finishes")
                            
                            # Count products and variants
                            unique_handles = feed_df['Handle'].unique()
                            
                            with st.container():
                                st.markdown('<div class="highlight">', unsafe_allow_html=True)
                                st.write(f"📊 Generated {len(unique_handles)} unique products with {len(feed_df)} total variants")
                                
                                # Create metrics
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Products", len(unique_handles))
                                with col2:
                                    st.metric("Variants", len(feed_df))
                                with col3:
                                    st.metric("Input Rows", len(valid_rows))
                                
                                st.markdown('</div>', unsafe_allow_html=True)
                            
                            # Preview of the generated feed
                            st.write("### Preview of Shopify Feed")
                            preview_cols = ['Handle', 'Title', 'Option1 Value', 'Option2 Value', 'Variant SKU', 'Variant Price']
                            st.dataframe(feed_df[preview_cols].head(20), use_container_width=True)
                            
                            # Download link
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            output_file = f"manual_shopify_feed_{timestamp}.xlsx"
                            st.markdown(get_excel_download_link(feed_df, output_file), unsafe_allow_html=True)
                        else:
                            st.error("❌ Failed to generate Shopify feed. The output is empty.")
                except Exception as e:
                    st.error(f"❌ Error generating Shopify feed: {e}")
                    st.write("Error details:", str(e))
        
        # Show example data
        with st.expander("📖 Example Data Structure", expanded=False):
            st.write("Here's how the Cadiz example would look in manual input:")
            example_data = [
                ["Cadiz Raised Circular Cupboard Knob", "32mm x 32mm p", "35607/1", "18.55", "##", "8"],
                ["Cadiz Raised Circular Cupboard Knob", "38mm x 32mm p", "35607/2", "20.58", "##", "8"],
            ]
            example_df = pd.DataFrame(example_data, columns=["Description", "Size", "SKU", "Price", "Finish Code", "Finish Count"])
            st.dataframe(example_df, use_container_width=True)
            
            st.write("**Finish Code Options:**")
            st.write("- `##` = Applies to 14 standard finishes")
            st.write("- `x##` = Applies to 8 premium finishes") 
            st.write("- Specific codes like `FFSB`, `FFPN` = Applies to that specific finish only")
            st.write("- Empty = Uses all available finishes")

if __name__ == "__main__":
    main() 