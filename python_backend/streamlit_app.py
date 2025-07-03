import streamlit as st
import pandas as pd
import pyodbc
import sys
import os
import tempfile
import zipfile
from datetime import datetime



sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from python_backend.fuzzy_matcher import FuzzyMatcher

def get_sql_server_connection():
    """
    Establish SQL Server connection using Windows Authentication (SSO)
    """
    try:
        conn = pyodbc.connect(
            "Driver={ODBC Driver 17 for SQL Server};"
            f"Server={st.session_state.server};"
            f"Database={st.session_state.database};"
            "Trusted_Connection=yes;"
        )
        return conn
    except Exception as e:
        st.error(f"Error connecting to SQL Server: {str(e)}")
        return None

def get_sql_tables(conn):
    """Get list of tables from the selected database"""
    cursor = conn.cursor()
    tables = cursor.tables(tableType='TABLE')
    return [table.table_name for table in tables]

def get_table_columns(conn, table_name):
    """Get columns from selected table"""
    query = f"SELECT TOP 1 * FROM {table_name}"
    df = pd.read_sql(query, conn)
    return df.columns.tolist()

# Set page config
st.set_page_config(
    page_title="Fuzzy Column Matcher",
    page_icon="🔍",
    layout="wide"
)

# Initialize session state
if 'matcher' not in st.session_state:
    st.session_state.matcher = FuzzyMatcher()

# Sidebar for configuration
with st.sidebar:
    st.header("Configuration")
    
    # Source type selection
    source_type = st.radio(
        "Select Source Type",
        ["Excel File", "SQL Server"]
    )
    
    # Target type selection
    target_type = st.radio(
        "Select Target Type",
        ["Excel File", "SQL Server", "Same Excel File"]
    )
    
    # Matching threshold
    threshold = st.slider(
        "Matching Threshold (%)",
        min_value=0,
        max_value=100,
        value=70,
        help="Minimum confidence score required for a match"
    )

# Title and description
st.title("🔍 Fuzzy Column Matcher")
st.markdown("""
This tool helps you match columns between Excel files and SQL Server tables using fuzzy logic.
It supports:
- Matching between different Excel files
- Matching between worksheets in the same Excel file
- Matching between Excel files and SQL Server tables
- Synonym-based matching
- Detailed matching results with confidence levels
""")

if source_type == "Excel File":
    st.info("💡 Tip: To match columns between worksheets in the same Excel file, select 'Same Excel File' as the Target Type after uploading your source Excel file.")

# Main content area
col1, col2 = st.columns(2)

# Source Data Selection
with col1:
    st.header("Source Data")
    source_df = None
    source_columns = []
    
    if source_type == "Excel File":
        source_file = st.file_uploader("Upload Source Excel File", type=['xlsx', 'xls'])
        if source_file:
            try:
                # Get list of worksheets
                source_bytes = source_file.getvalue()
                excel_file = pd.ExcelFile(pd.io.common.BytesIO(source_bytes))
                worksheets = excel_file.sheet_names
                if not worksheets:
                    st.error("No worksheets found in the source Excel file")
                    source_df = None
                    source_columns = []
                else:
                    # Let user select worksheet
                    selected_worksheet = st.selectbox(
                        "Select Source Worksheet",
                        options=worksheets,
                        key="source_worksheet"
                    )
                    
                    # Read the selected worksheet
                    try:
                        # Create DataFrame from the selected worksheet
                        df = pd.read_excel(pd.io.common.BytesIO(source_bytes), sheet_name=selected_worksheet)
                        
                        # Validate DataFrame
                        if df is None or df.empty:
                            st.error("Selected worksheet appears to be empty")
                            source_df = None
                            source_columns = []
                        else:
                            # Convert to proper DataFrame if needed
                            if not isinstance(df, pd.DataFrame):
                                df = pd.DataFrame(df)
                            
                            # Store in session state
                            source_df = df.copy()
                            st.session_state.source_df = source_df
                            source_columns = source_df.columns.tolist()
                            st.write("Source worksheet loaded successfully")
                            st.write(f"Available columns: {source_columns}")
                            
                            # Only show column selector if we have columns
                            if source_columns:
                                st.session_state.source_column = st.selectbox(
                                    "Select Source Column",
                                    options=source_columns
                                )
                    except pd.errors.EmptyDataError:
                        st.error("The selected worksheet is empty")
                        source_df = None
                        source_columns = []
                    except Exception as e:
                        st.error(f"Error reading worksheet: {str(e)}")
                        source_df = None
                        source_columns = []
            except zipfile.BadZipFile:
                st.error("The uploaded file is not a valid Excel file. Please ensure you're uploading a valid .xlsx or .xls file.")
                source_df = None
                source_columns = []
            except Exception as e:
                st.error(f"Error reading source Excel file: {str(e)}")
                source_df = None
                source_columns = []
    else:  # SQL Server
        st.session_state.server = st.text_input("SQL Server Name")
        st.session_state.database = st.text_input("Database Name")
        
        if st.session_state.server and st.session_state.database:
            conn = get_sql_server_connection()
            if conn:
                tables = get_sql_tables(conn)
                selected_table = st.selectbox("Select Table", tables)
                if selected_table:
                    columns = get_table_columns(conn, selected_table)
                    st.session_state.source_column = st.selectbox(
                        "Select Source Column",
                        options=columns
                    )
                    source_df = pd.read_sql(f"SELECT * FROM {selected_table}", conn)
                    st.session_state.source_df = source_df

# Target Data Selection
with col2:
    st.header("Target Data")
    target_df = None
    target_columns = []
    
    if target_type == "Same Excel File" and source_file:
        # Use the same Excel file as source
        try:
            # Get list of worksheets (excluding the source worksheet)
            excel_file = pd.ExcelFile(pd.io.common.BytesIO(st.session_state.source_bytes))
            worksheets = [ws for ws in excel_file.sheet_names if ws != st.session_state.get('source_worksheet')]
            if not worksheets:
                st.error("No additional worksheets found in the Excel file")
                target_df = None
                target_columns = []
            else:
                # Let user select worksheet
                selected_worksheet = st.selectbox(
                    "Select Target Worksheet",
                    options=worksheets,
                    key="target_worksheet"
                )
                
                # Read the selected worksheet
                try:
                    # Create DataFrame from the selected worksheet
                    df = pd.read_excel(pd.io.common.BytesIO(st.session_state.source_bytes), sheet_name=selected_worksheet)
                    
                    # Validate DataFrame
                    if df is None or df.empty:
                        st.error("Selected worksheet appears to be empty")
                        target_df = None
                        target_columns = []
                    else:
                        # Convert to proper DataFrame if needed
                        if not isinstance(df, pd.DataFrame):
                            df = pd.DataFrame(df)
                        
                        # Store in session state
                        target_df = df.copy()
                        st.session_state.target_df = target_df
                        target_columns = target_df.columns.tolist()
                        st.write("Target worksheet loaded successfully")
                        st.write(f"Available columns: {target_columns}")
                        
                        # Only show column selector if we have columns
                        if target_columns:
                            st.session_state.target_column = st.selectbox(
                                "Select Target Column",
                                options=target_columns
                            )
                except pd.errors.EmptyDataError:
                    st.error("The selected worksheet is empty")
                    target_df = None
                    target_columns = []
                except Exception as e:
                    st.error(f"Error reading worksheet: {str(e)}")
                    target_df = None
                    target_columns = []
        except Exception as e:
            st.error(f"Error reading target worksheet: {str(e)}")
            target_df = None
            target_columns = []
    
    elif target_type == "Excel File":
        target_file = st.file_uploader("Upload Target Excel File", type=['xlsx', 'xls'])
        if target_file:
            try:
                # Get list of worksheets
                target_bytes = target_file.getvalue()
                excel_file = pd.ExcelFile(pd.io.common.BytesIO(target_bytes))
                worksheets = excel_file.sheet_names
                if not worksheets:
                    st.error("No worksheets found in the target Excel file")
                    target_df = None
                    target_columns = []
                else:
                    # Let user select worksheet
                    selected_worksheet = st.selectbox(
                        "Select Target Worksheet",
                        options=worksheets,
                        key="target_worksheet"
                    )
                    
                    # Read the selected worksheet
                    try:
                        # Create DataFrame from the selected worksheet
                        df = pd.read_excel(pd.io.common.BytesIO(target_bytes), sheet_name=selected_worksheet)
                        
                        # Validate DataFrame
                        if df is None or df.empty:
                            st.error("Selected worksheet appears to be empty")
                            target_df = None
                            target_columns = []
                        else:
                            # Convert to proper DataFrame if needed
                            if not isinstance(df, pd.DataFrame):
                                df = pd.DataFrame(df)
                            
                            # Store in session state
                            target_df = df.copy()
                            st.session_state.target_df = target_df
                            target_columns = target_df.columns.tolist()
                            st.write("Target worksheet loaded successfully")
                            st.write(f"Available columns: {target_columns}")
                            
                            # Only show column selector if we have columns
                            if target_columns:
                                st.session_state.target_column = st.selectbox(
                                    "Select Target Column",
                                    options=target_columns
                                )
                    except Exception as e:
                        st.error(f"Error reading worksheet: {str(e)}")
                        target_df = None
                        target_columns = []
            except zipfile.BadZipFile:
                st.error("The uploaded file is not a valid Excel file. Please ensure you're uploading a valid .xlsx or .xls file.")
                target_df = None
                target_columns = []
            except Exception as e:
                st.error(f"Error reading target Excel file: {str(e)}")
                target_df = None
                target_columns = []
    
    else:  # SQL Server
        if 'server' not in st.session_state:
            st.session_state.server = st.text_input("SQL Server Name ", key="target_server")
            st.session_state.database = st.text_input("Database Name ", key="target_db")
        
        if st.session_state.server and st.session_state.database:
            conn = get_sql_server_connection()
            if conn:
                tables = get_sql_tables(conn)
                selected_table = st.selectbox("Select Table ", tables)
                if selected_table:
                    columns = get_table_columns(conn, selected_table)
                    st.session_state.target_column = st.selectbox(
                        "Select Target Column",
                        options=columns
                    )
                    target_df = pd.read_sql(f"SELECT * FROM {selected_table}", conn)
                    st.session_state.target_df = target_df

# Process matching
if st.button("Run Matching"):
    if (hasattr(st.session_state, 'source_df') and hasattr(st.session_state, 'target_df') and 
        st.session_state.source_df is not None and st.session_state.target_df is not None):
        with st.spinner("Processing matches..."):
            # Update matcher threshold
            st.session_state.matcher.threshold = threshold
            
            success = True
            results = None
            results_df = None

            try:
                # Get DataFrames from session state
                source_df = st.session_state.source_df
                target_df = st.session_state.target_df

                # Debug information
                st.write("Validating DataFrames...")
                
                # Validate source DataFrame
                if not isinstance(source_df, pd.DataFrame):
                    raise ValueError("Source data is not a valid DataFrame")
                if source_df.empty:
                    raise ValueError("Source DataFrame is empty")
                if not hasattr(st.session_state, 'source_column'):
                    raise ValueError("Source column not selected")
                if st.session_state.source_column not in source_df.columns:
                    raise ValueError(f"Source column '{st.session_state.source_column}' not found")
                
                st.write("Source DataFrame validated successfully")
                st.write(f"Shape: {source_df.shape}")
                st.write(f"Selected column: {st.session_state.source_column}")
                
                # Validate target DataFrame
                if not isinstance(target_df, pd.DataFrame):
                    raise ValueError("Target data is not a valid DataFrame")
                if target_df.empty:
                    raise ValueError("Target DataFrame is empty")
                if not hasattr(st.session_state, 'target_column'):
                    raise ValueError("Target column not selected")
                if st.session_state.target_column not in target_df.columns:
                    raise ValueError(f"Target column '{st.session_state.target_column}' not found")
                
                st.write("Target DataFrame validated successfully")
                st.write(f"Shape: {target_df.shape}")
                st.write(f"Selected column: {st.session_state.target_column}")
                
                st.write("\nStarting matching process...")
                # Create fresh copies to avoid modifying originals
                source_copy = source_df.copy()
                target_copy = target_df.copy()
                
                try:
                    # Perform matching
                    st.write("Running matching algorithm...")
                    results = st.session_state.matcher.match_columns(
                        source_copy,
                        target_copy,
                        st.session_state.source_column,
                        st.session_state.target_column
                    )
                    st.write("✅ Matching process completed successfully")
                except Exception as e:
                    st.error(f"Error during matching: {str(e)}")
                    st.error("Please ensure your data is in the correct format")
                    success = False
            except ValueError as ve:
                st.error(f"Validation Error: {str(ve)}")
                success = False
            except pd.errors.EmptyDataError:
                st.error("One or both DataFrames are empty. Please check your data.")
                success = False
            except Exception as e:
                st.error(f"Unexpected error: {str(e)}")
                st.error("Please ensure your data is valid and try again.")
                success = False

            if success and results is not None:
                try:
                    st.write("Formatting results...")
                    # Format results
                    results_df = st.session_state.matcher.format_results_for_export(results)
                    if results_df.empty:
                        st.warning("No matches found between the selected columns")
                        success = False
                    else:
                        st.write("✅ Results formatted successfully")
                except pd.errors.EmptyDataError:
                    st.error("No data to format. The matching process returned empty results.")
                    success = False
                except Exception as e:
                    st.error(f"Error formatting results: {str(e)}")
                    st.error("There was an issue processing the matching results. Please check your data.")
                    success = False

            if success and results_df is not None:
                # Display results
                st.header("Matching Results")
                
                # Display summary metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Matches", len(results["matches"]))
                with col2:
                    st.metric("Source Mismatches", len(results["source_mismatches"]))
                with col3:
                    st.metric("Target Mismatches", len(results["target_mismatches"]))
                
                # Display detailed results
                st.dataframe(results_df)
                
                # Prepare results for download
                try:
                    st.write("Preparing download file...")
                    # Create BytesIO object
                    output = pd.io.common.BytesIO()
                    
                    # Write to BytesIO buffer
                    with pd.ExcelWriter(output, engine='openpyxl', mode='w') as writer:
                        # Write main results
                        results_df.to_excel(writer, sheet_name='Matching Results', index=False)
                        
                        # Create summary data
                        summary_data = pd.DataFrame({
                            'Metric': [
                                'Total Records Processed',
                                'Successful Matches',
                                'Source Mismatches',
                                'Target Mismatches',
                                'Average Confidence Score'
                            ],
                            'Value': [
                                len(results_df),
                                len(results_df[results_df['Type'] == 'Match']),
                                len(results_df[results_df['Type'] == 'Source Mismatch']),
                                len(results_df[results_df['Type'] == 'Target Mismatch']),
                                f"{results_df['Confidence'].str.rstrip('%').astype(float).mean():.2f}%"
                            ]
                        })
                        
                        # Write summary sheet
                        summary_data.to_excel(writer, sheet_name='Summary', index=False)
                    
                    # Get the value of the BytesIO buffer
                    excel_data = output.getvalue()
                    
                    st.write("File prepared successfully. Ready for download.")
                    # Create download button
                    st.download_button(
                        label="📥 Download Results",
                        data=excel_data,
                        file_name=f"fuzzy_matching_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Download the matching results as an Excel file"
                    )
                except Exception as e:
                    st.error(f"Could not prepare the download: {str(e)}")
                    st.error("Please try the matching process again.")
    else:
        st.error("Please select both source and target data before running the matching process.")

# Footer
st.markdown("---")
st.markdown("""
### Tips
- For better matching results, try adjusting the threshold value in the sidebar
- The confidence score indicates how well the columns match
- Download the results to get a detailed Excel report including matches and mismatches
""")
