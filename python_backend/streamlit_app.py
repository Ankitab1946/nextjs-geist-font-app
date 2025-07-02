import streamlit as st
import pandas as pd
import pyodbc
import sys
import os
import tempfile
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

def save_results(results_df):
    """Save results to Excel file and provide download link"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"fuzzy_matching_results_{timestamp}.xlsx"
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write main results
        results_df.to_excel(writer, sheet_name='Matching Results', index=False)
        
        # Add summary sheet
        summary_data = {
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
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
        
    return filename

# Set page config
st.set_page_config(
    page_title="Fuzzy Column Matcher",
    page_icon="🔍",
    layout="wide"
)

# Initialize session state
if 'matcher' not in st.session_state:
    st.session_state.matcher = FuzzyMatcher()

# Title and description
st.title("🔍 Fuzzy Column Matcher")
st.markdown("""
This tool helps you match columns between Excel files and SQL Server tables using fuzzy logic.
It supports synonyms and provides detailed matching results with confidence levels.
""")

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
        ["Excel File", "SQL Server"]
    )
    
    # Matching threshold
    threshold = st.slider(
        "Matching Threshold (%)",
        min_value=0,
        max_value=100,
        value=70,
        help="Minimum confidence score required for a match"
    )

# Main content area
col1, col2 = st.columns(2)

# Source Data Selection
with col1:
    st.header("Source Data")
    
    if source_type == "Excel File":
        source_file = st.file_uploader("Upload Source Excel File", type=['xlsx', 'xls'])
        if source_file:
            source_df = pd.read_excel(source_file)
            st.session_state.source_df = source_df
            source_columns = source_df.columns.tolist()
            st.session_state.source_column = st.selectbox(
                "Select Source Column",
                options=source_columns
            )
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
    
    if target_type == "Excel File":
        target_file = st.file_uploader("Upload Target Excel File", type=['xlsx', 'xls'])
        if target_file:
            target_df = pd.read_excel(target_file)
            st.session_state.target_df = target_df
            target_columns = target_df.columns.tolist()
            st.session_state.target_column = st.selectbox(
                "Select Target Column",
                options=target_columns
            )
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
    if hasattr(st.session_state, 'source_df') and hasattr(st.session_state, 'target_df'):
        with st.spinner("Processing matches..."):
            # Update matcher threshold
            st.session_state.matcher.threshold = threshold
            
            # Perform matching
            results = st.session_state.matcher.match_columns(
                st.session_state.source_df,
                st.session_state.target_df,
                st.session_state.source_column,
                st.session_state.target_column
            )
            
            # Format results
            results_df = st.session_state.matcher.format_results_for_export(results)
            
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
            
            # Save and provide download link
            filename = save_results(results_df)
            with open(filename, "rb") as file:
                st.download_button(
                    label="Download Results",
                    data=file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Cleanup temporary file
            os.remove(filename)
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
