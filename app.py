import streamlit as st
import pandas as pd
from io import BytesIO
import base64
from pathlib import Path

def get_download_link(df, filename, sheet_name='Sheet1'):
    """Generate a download link for the DataFrame as an Excel file"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=isinstance(df.index, pd.MultiIndex))
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'

def pivot_and_merge_files(files):
    """Pivot each Excel file and then merge them into a single DataFrame."""
    all_pivoted_dfs = []
    
    for file in files:
        try:
            df = pd.read_excel(file, engine='openpyxl')
            
            # Check if required columns exist
            required_columns = {'CID', 'Month', 'Amount'}
            if not required_columns.issubset(df.columns):
                st.error(f"File '{file.name}' is missing required columns. Needs: CID, Month, Amount")
                continue
                
            # Pivot the data
            pivoted_df = df.pivot_table(
                index='CID',
                columns='Month',
                values='Amount',
                aggfunc='sum',
                fill_value=0
            )
            
            # Ensure all month columns exist (APR, MAY, JUN or APR-25, MAY-25, JUN-25)
            month_columns = []
            for month in ['APR', 'MAY', 'JUN']:
                col_with_year = f"{month}-25"
                if col_with_year in pivoted_df.columns:
                    month_columns.append(col_with_year)
                elif month in pivoted_df.columns:
                    month_columns.append(month)
                else:
                    pivoted_df[col_with_year] = 0
                    month_columns.append(col_with_year)
            
            # Select and reorder columns
            pivoted_df = pivoted_df[month_columns]
            pivoted_df = pivoted_df.reset_index()
            pivoted_df['Source_File'] = file.name
            all_pivoted_dfs.append(pivoted_df)
            
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
            return None
    
    if not all_pivoted_dfs:
        return None
        
    # Concatenate all pivoted DataFrames
    merged_df = pd.concat(all_pivoted_dfs, ignore_index=True)
    
    # Group by CID and sum the values
    numeric_cols = [col for col in merged_df.columns if col not in ['CID', 'Source_File']]
    merged_df = merged_df.groupby('CID', as_index=False)[numeric_cols].sum()
    
    # Rename CID to SCNO
    merged_df = merged_df.rename(columns={'CID': 'SCNO'})
    
    # Reorder columns to put SCNO first, then months
    final_columns = ['SCNO'] + [col for col in merged_df.columns if col != 'SCNO']
    return merged_df[final_columns]

def pivot_data(df):
    """Pivot the data with CID as rows and Month as columns."""
    try:
        required_columns = {'CID', 'Month', 'Amount'}
        if not required_columns.issubset(df.columns):
            st.error("Missing required columns. Needs: CID, Month, Amount")
            return None
        
        pivoted_df = df.pivot_table(
            index='CID',
            columns='Month',
            values='Amount',
            aggfunc='sum',
            fill_value=0
        )
        
        month_order = ['APR-25', 'MAY-25', 'JUN-25']
        existing_months = [m for m in month_order if m in pivoted_df.columns]
        if existing_months:
            pivoted_df = pivoted_df[existing_months]
            
        return pivoted_df
    except Exception as e:
        st.error(f"Error pivoting data: {str(e)}")
        return None

def simple_merge_files(files):
    """Merge Excel files that are already in CID + month columns format."""
    all_dfs = []
    month_order = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    
    for file in files:
        try:
            df = pd.read_excel(file, engine='openpyxl')
            if 'CID' not in df.columns:
                st.error(f"File '{file.name}' is missing the 'CID' column.")
                continue
            df['Source_File'] = file.name
            all_dfs.append(df)
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
            continue
    
    if not all_dfs:
        return None
    
    merged_df = pd.concat(all_dfs, ignore_index=True)
    value_columns = [col for col in merged_df.columns if col not in ['CID', 'Source_File']]
    merged_df = merged_df.groupby('CID', as_index=False)[value_columns].sum()
    merged_df = merged_df.rename(columns={'CID': 'SCNO'})
    
    month_columns = []
    for month in month_order:
        cols = [col for col in merged_df.columns 
                if col != 'SCNO' and any(m in str(col).upper() for m in [month, f"{month}-"])]
        cols_sorted = sorted(cols, key=lambda x: int(x.split('-')[1]) if '-' in str(x) and len(str(x).split('-')) > 1 and str(x).split('-')[1].isdigit() else 0)
        month_columns.extend(cols_sorted)
    
    other_cols = [col for col in merged_df.columns if col not in ['SCNO'] + month_columns]
    final_columns = ['SCNO'] + month_columns + other_cols
    return merged_df[final_columns]

def merge_format3_files(files):
    """Merge Excel files in Format 3 (CID + month columns)."""
    all_dfs = []
    
    for file in files:
        try:
            df = pd.read_excel(file, engine='openpyxl')
            if 'CID' not in df.columns:
                st.error(f"File '{file.name}' is missing the 'CID' column.")
                continue
            df['Source_File'] = file.name
            all_dfs.append(df)
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
            continue
    
    if not all_dfs:
        return None
    
    merged_df = pd.concat(all_dfs, ignore_index=True)
    value_columns = [col for col in merged_df.columns if col not in ['CID', 'Source_File']]
    merged_df = merged_df.groupby('CID', as_index=False)[value_columns].sum()
    merged_df = merged_df.rename(columns={'CID': 'SCNO'})
    
    return merged_df

def final_merge_tab():
    """Content for the Final Merge tab."""
    st.header("Final Merge: Match and Calculate")
    st.write("Merge two Excel files by SCNO and calculate highest bill")
    
    # Add custom CSS to hide the file uploader's max size warning
    st.markdown("""
    <style>
    .stAlert {display: none;}
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # File uploaders with size limit info
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("File 1: Template")
        st.write("Expected columns: SCNO NO, APR-25, MAY-25, JUNE-25, Highest Bill")
        st.write("Max file size: 200MB")
        file1 = st.file_uploader(
            "Upload Template File", 
            type=["xlsx"], 
            key="tab5_final_merge_file1",
            help="Upload your template file with SCNO and empty month columns"
        )
    
    with col2:
        st.subheader("File 2: Data")
        st.write("Expected columns: SCNO NO, APR-25, MAY-25, JUNE-25")
        st.write("Max file size: 200MB")
        file2 = st.file_uploader(
            "Upload Data File", 
            type=["xlsx"], 
            key="tab5_final_merge_file2",
            help="Upload your data file with SCNO and month values"
        )
    
    # Add file size validation
    MAX_FILE_SIZE = 200 * 1024 * 1024  # 200MB in bytes
    
    def validate_file_size(file):
        if file.size > MAX_FILE_SIZE:
            st.error(f"File {file.name} is too large. Maximum size is 200MB.")
            return False
        return True
    
    if file1 and file2:
        # Validate file sizes
        if not all(validate_file_size(f) for f in [file1, file2]):
            return
            
        if st.button("Merge and Process", key="tab5_merge_process_button", type="primary"):
            with st.spinner("Processing files..."):
                try:
                    # Create progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Read files with optimized parameters for large files
                    status_text.text("Reading files...")
                    progress_bar.progress(10)
                    
                    # Read in chunks for large files
                    chunksize = 10000  # Adjust based on your needs
                    
                    # Read template file
                    df_template = pd.read_excel(
                        file1, 
                        engine='openpyxl',
                        dtype=str  # Read everything as string first to preserve formatting
                    )
                    
                    # Read data file in chunks if large
                    if file2.size > 10 * 1024 * 1024:  # If file is larger than 10MB
                        chunks = []
                        for chunk in pd.read_excel(
                            file2, 
                            engine='openpyxl',
                            chunksize=chunksize,
                            dtype=str
                        ):
                            chunks.append(chunk)
                        df_data = pd.concat(chunks, ignore_index=True)
                    else:
                        df_data = pd.read_excel(file2, engine='openpyxl', dtype=str)
                    
                    # Convert numeric columns to float
                    for col in ['APR-25', 'MAY-25', 'JUNE-25']:
                        if col in df_data.columns:
                            df_data[col] = pd.to_numeric(df_data[col], errors='coerce').fillna(0)
                    
                    # Validate columns
                    required_cols = ['SCNO NO', 'APR-25', 'MAY-25', 'JUNE-25']
                    for col in required_cols:
                        if col not in df_template.columns:
                            st.error(f"Template file is missing required column: {col}")
                            return
                        if col not in df_data.columns:
                            st.error(f"Data file is missing required column: {col}")
                            return
                    
                    # Ensure SCNO NO is string type for proper matching
                    df_template['SCNO NO'] = df_template['SCNO NO'].astype(str).str.strip()
                    df_data['SCNO NO'] = df_data['SCNO NO'].astype(str).str.strip()
                    
                    status_text.text("Matching SCNO numbers...")
                    progress_bar.progress(30)
                    
                    # Create a dictionary for faster lookup
                    data_dict = df_data.set_index('SCNO NO').to_dict('index')
                    
                    # Initialize new columns if they don't exist
                    if 'Highest Bill' not in df_template.columns:
                        df_template['Highest Bill'] = 0
                    if 'Match Status' not in df_template.columns:
                        df_template['Match Status'] = ''
                    
                    status_text.text("Processing matches...")
                    progress_bar.progress(50)
                    
                    # Process each row in the template
                    for idx, row in df_template.iterrows():
                        scno = str(row['SCNO NO']).strip()
                        
                        if scno in data_dict:
                            # Update month values
                            for month in ['APR-25', 'MAY-25', 'JUNE-25']:
                                df_template.at[idx, month] = data_dict[scno].get(month, 0)
                            
                            # Calculate highest bill
                            values = [data_dict[scno].get(month, 0) for month in ['APR-25', 'MAY-25', 'JUNE-25']]
                            df_template.at[idx, 'Highest Bill'] = max(values)
                            df_template.at[idx, 'Match Status'] = 'M'
                        else:
                            # No match found
                            for month in ['APR-25', 'MAY-25', 'JUNE-25']:
                                df_template.at[idx, month] = ''
                            df_template.at[idx, 'Highest Bill'] = ''
                            df_template.at[idx, 'Match Status'] = 'N/A'
                    
                    progress_bar.progress(90)
                    status_text.text("Finalizing results...")
                    
                    # Reorder columns
                    final_columns = ['SCNO NO', 'APR-25', 'MAY-25', 'JUNE-25', 'Highest Bill', 'Match Status']
                    df_result = df_template[final_columns]
                    
                    # Store result in session state
                    st.session_state.final_merge_result = df_result
                    
                    progress_bar.progress(100)
                    status_text.text("Processing complete!")
                    
                    # Show success message and preview
                    st.success("Files processed successfully!")
                    
                    # Show summary
                    total_rows = len(df_result)
                    matched = len(df_result[df_result['Match Status'] == 'M'])
                    not_matched = len(df_result[df_result['Match Status'] == 'N/A'])
                    
                    st.write("### Summary")
                    st.write(f"- Total Rows: {total_rows}")
                    st.write(f"- Matched SCNO: {matched}")
                    st.write(f"- Not Matched: {not_matched}")
                    
                    # Show preview
                    st.write("### Preview of Results")
                    st.dataframe(df_result.head())
                    
                    # Download button
                    st.markdown("### Download Results")
                    st.markdown(get_download_link(df_result, "final_merge_result.xlsx", "Download Merged Results"), 
                                unsafe_allow_html=True)
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())
                finally:
                    progress_bar.empty()

def final_merge_single_file_tab():
    """Content for the Final Merge Single File tab."""
    st.header("Final Merge: Single File Processing")
    st.write("Process large Excel files with Main Sheet and Value Sheet for SCNO matching")
    
    # Add custom CSS for progress bar
    st.markdown("""
    <style>
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # File uploader for the Excel file
    uploaded_file = st.file_uploader(
        "Upload Excel file with Main Sheet and Value Sheet",
        type=["xlsx"],
        key="tab6_single_file_uploader",
        help="Upload an Excel file with two sheets: 'Main Sheet' and 'Value Sheet'"
    )
    
    if uploaded_file is not None:
        if st.button("Process File", key="tab6_process_button", type="primary"):
            with st.spinner("Processing file..."):
                try:
                    # Initialize progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Read the Excel file
                    status_text.text("Reading Excel file...")
                    progress_bar.progress(10)
                    
                    # Read both sheets
                    xls = pd.ExcelFile(uploaded_file)
                    
                    # Check if required sheets exist
                    if 'Main Sheet' not in xls.sheet_names or 'Value Sheet' not in xls.sheet_names:
                        st.error("Error: The Excel file must contain both 'Main Sheet' and 'Value Sheet'")
                        return
                    
                    # Read Main Sheet
                    status_text.text("Reading Main Sheet...")
                    main_sheet = pd.read_excel(
                        xls, 
                        sheet_name='Main Sheet',
                        dtype={'SCNO': str}
                    )
                    
                    # Ensure required columns exist in Main Sheet
                    required_main_cols = ['SCNO', 'APR-25', 'MAY-25', 'JUNE-25', 'Highest Bill']
                    for col in required_main_cols:
                        if col not in main_sheet.columns:
                            main_sheet[col] = ''  # Add missing columns
                    
                    # Read Value Sheet
                    status_text.text("Reading Value Sheet...")
                    value_sheet = pd.read_excel(
                        xls, 
                        sheet_name='Value Sheet',
                        dtype={'SCNO': str}
                    )
                    
                    # Ensure required columns exist in Value Sheet
                    required_value_cols = ['SCNO', 'APR-25', 'MAY-25', 'JUNE-25']
                    for col in required_value_cols:
                        if col not in value_sheet.columns:
                            st.error(f"Error: Column '{col}' is missing in Value Sheet")
                            return
                    
                    # Convert SCNO to string and strip whitespace
                    main_sheet['SCNO'] = main_sheet['SCNO'].astype(str).str.strip()
                    value_sheet['SCNO'] = value_sheet['SCNO'].astype(str).str.strip()
                    
                    # Create a dictionary for faster lookup
                    status_text.text("Preparing data for processing...")
                    progress_bar.progress(30)
                    
                    value_dict = value_sheet.set_index('SCNO').to_dict('index')
                    
                    # Process each row in main sheet
                    status_text.text("Matching SCNO numbers...")
                    total_rows = len(main_sheet)
                    processed_rows = 0
                    
                    # Initialize progress tracking
                    progress_step = max(1, total_rows // 100)  # Update progress every 1%
                    
                    for idx, row in main_sheet.iterrows():
                        scno = str(row['SCNO']).strip()
                        
                        if scno in value_dict:
                            # Update month values
                            for month in ['APR-25', 'MAY-25', 'JUNE-25']:
                                main_sheet.at[idx, month] = value_dict[scno].get(month, '')
                            
                            # Calculate highest bill
                            values = [
                                value_dict[scno].get('APR-25', 0),
                                value_dict[scno].get('MAY-25', 0),
                                value_dict[scno].get('JUNE-25', 0)
                            ]
                            # Convert to numeric, handle empty strings
                            numeric_values = []
                            for v in values:
                                try:
                                    numeric_values.append(float(v) if str(v).strip() else 0)
                                except (ValueError, TypeError):
                                    numeric_values.append(0)
                            
                            if any(numeric_values):
                                main_sheet.at[idx, 'Highest Bill'] = max(numeric_values)
                        
                        # Update progress
                        processed_rows += 1
                        if processed_rows % progress_step == 0 or processed_rows == total_rows:
                            progress = int(30 + (70 * processed_rows / total_rows))
                            progress_bar.progress(min(progress, 100))
                    
                    # Final progress update
                    progress_bar.progress(100)
                    status_text.text("Processing complete!")
                    
                    # Show success message and summary
                    st.success("File processed successfully!")
                    
                    # Show preview
                    st.write("### Preview of Processed Data")
                    st.dataframe(main_sheet.head())
                    
                    # Add download button
                    st.markdown("### Download Processed File")
                    st.markdown(
                        get_download_link(main_sheet, "processed_final_merge.xlsx", "Download Processed File"),
                        unsafe_allow_html=True
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    import traceback
                    st.error(traceback.format_exc())
                finally:
                    progress_bar.empty()

def main():
    st.set_page_config(
        page_title="Excel Tools",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel Tools")
    
    # Create tabs for different functionalities
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Merge and Pivot", 
        "Pivot Data", 
        "Simple Merge", 
        "Format 3 Merge", 
        "Final Merge",
        "Final Merge Single File"
    ])
    
    with tab1:
        st.header("Merge and Pivot Excel Files")
        st.write("Upload Excel (.xlsx) files to pivot and merge them into a single file.")
        st.write("Each file should contain: CID, Month, and Amount columns.")
        
        merge_files = st.file_uploader(
            "Choose Excel files to merge",
            type=['xlsx'],
            accept_multiple_files=True,
            key="tab1_merge_uploader"
        )
        
        if merge_files:
            if len(merge_files) > 10:
                st.warning("You can upload a maximum of 10 files. Only the first 10 will be processed.")
                merge_files = merge_files[:10]
            
            st.write(f"### Files to merge ({len(merge_files)}):")
            for file in merge_files:
                st.write(f"- {file.name}")
            
            if st.button("Process and Merge Files", key="merge_pivot_button"):
                with st.spinner("Processing and merging files..."):
                    merged_df = pivot_and_merge_files(merge_files)
                    if merged_df is not None:
                        st.success("Files processed and merged successfully!")
                        st.write("### Preview of Merged Data")
                        st.dataframe(merged_df.head())
                        st.write("### Data Summary")
                        st.write(f"- Total Rows: {len(merged_df)}")
                        st.write(f"- Total Columns: {len(merged_df.columns)}")
                        st.write(f"- Columns: {', '.join(merged_df.columns)}")
                        st.markdown("### Download Merged File")
                        st.markdown(
                            get_download_link(merged_df, "merged_pivoted_data.xlsx", "Download Merged Data"), 
                            unsafe_allow_html=True
                        )
    
    with tab2:
        st.header("Pivot Excel Data")
        st.write("Upload Excel files to pivot CID vs Month data")
        
        pivot_files = st.file_uploader(
            "Choose Excel files to pivot",
            type=["xlsx"],
            accept_multiple_files=True,
            key="tab2_pivot_uploader"
        )
        
        if pivot_files:
            for uploaded_file in pivot_files:
                try:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                    pivoted_df = pivot_data(df)
                    
                    if pivoted_df is not None:
                        st.subheader(f"Pivoted Data: {uploaded_file.name}")
                        st.dataframe(pivoted_df)
                        st.write("### Data Summary")
                        st.write(f"- Total Rows: {len(pivoted_df)}")
                        st.write(f"- Total Columns: {len(pivoted_df.columns)}")
                        st.write(f"- Columns: {', '.join(pivoted_df.columns)}")
                        
                        output_filename = f"pivoted_{uploaded_file.name}"
                        st.markdown("### Download Pivoted File")
                        st.markdown(
                            get_download_link(pivoted_df, output_filename, "Download Pivoted Data"), 
                            unsafe_allow_html=True
                        )
                        st.write("---")
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {str(e)}")
    
    with tab3:
        st.header("Simple Merge")
        st.write("Upload Excel (.xlsx) files with the same structure to merge")
        st.write("Each file should contain: CID, Month, and Amount columns.")
        
        simple_merge_files_upload = st.file_uploader(
            "Choose Excel files to merge",
            type=["xlsx"],
            accept_multiple_files=True,
            key="tab3_simple_merge_uploader"
        )
        
        if simple_merge_files_upload:
            st.write(f"### Files to merge ({len(simple_merge_files_upload)}):")
            for file in simple_merge_files_upload:
                st.write(f"- {file.name}")
            
            if st.button("Merge Files", key="simple_merge_button"):
                with st.spinner("Merging files..."):
                    merged_df = simple_merge_files(simple_merge_files_upload)
                    if merged_df is not None:
                        st.success("Files merged successfully!")
                        st.write("### Preview of Merged Data")
                        st.dataframe(merged_df.head())
                        st.write("### Data Summary")
                        st.write(f"- Total Rows: {len(merged_df)}")
                        st.write(f"- Total Columns: {len(merged_df.columns)}")
                        st.write(f"- Columns: {', '.join(merged_df.columns)}")
                        st.markdown("### Download Merged File")
                        st.markdown(
                            get_download_link(merged_df, "merged_data.xlsx", "Download Merged Data"),
                            unsafe_allow_html=True
                        )
    
    with tab4:
        st.header("Format 3 Merge")
        st.write("Upload Excel files in Format 3 (CID + month columns)")
        st.write("Example: CID as first column, followed by month columns (e.g., MAY-25, APR-25)")
        
        format3_files = st.file_uploader(
            "Choose Excel files to merge",
            type=["xlsx"],
            accept_multiple_files=True,
            key="tab4_format3_merge_uploader"
        )
        
        if format3_files:
            st.write(f"### Files to merge ({len(format3_files)}):")
            for file in format3_files:
                st.write(f"- {file.name}")
            
            if st.button("Merge Files", key="format3_merge_button"):
                if 'merged_data' in st.session_state:
                    del st.session_state.merged_data
                
                with st.spinner("Merging files..."):
                    merged_df = merge_format3_files(format3_files)
                    
                    if merged_df is not None:
                        st.session_state.merged_data = merged_df
                        st.success("✅ Files merged successfully!")
                        st.write("### Preview of Merged Data")
                        st.dataframe(merged_df.head())
                        st.write("### Data Summary")
                        st.write(f"- Total Rows: {len(merged_df)}")
                        st.write(f"- Total Columns: {len(merged_df.columns)}")
                        st.write(f"- Columns: {', '.join(merged_df.columns)}")
                        st.markdown("### Download Merged File")
                        st.markdown(
                            get_download_link(merged_df, "format3_merged_output.xlsx", "Download Merged Data"),
                            unsafe_allow_html=True
                        )
    
    with tab5:
        final_merge_tab()
        
    with tab6:
        final_merge_single_file_tab()

if __name__ == "__main__":
    main()
