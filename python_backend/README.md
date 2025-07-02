# Fuzzy Column Matcher

A Python-based framework for comparing columns using fuzzy logic, supporting Excel and SQL Server data sources with SSO authentication.

## Features

- 🔍 Fuzzy matching with synonym support
- 📊 Excel-to-Excel and Excel-to-SQL comparisons
- 🔐 SQL Server SSO authentication
- ⚡ Support for large datasets using Dask
- 📈 Two-way matching with confidence scores
- 📥 Downloadable detailed results
- 🎯 Customizable matching threshold
- 🔄 Synonym support for business terms

## Prerequisites

- Python 3.8+
- SQL Server ODBC Driver 17 for SQL Server connections
- Windows Authentication configured for SSO (when using SQL Server)

## Installation

1. Clone the repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

The application uses `rapidfuzz` for efficient fuzzy string matching, which is easier to install and more performant than alternatives.

3. Install SQL Server ODBC Driver if not already installed:
   - [Download SQL Server ODBC Driver](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)

## Usage

### Running the Streamlit App

```bash
streamlit run streamlit_app.py
```

This will launch the web interface where you can:
1. Select source and target data sources (Excel/SQL)
2. Choose columns to compare
3. Adjust matching threshold
4. Run the comparison
5. Download results

### Using the Python API

```python
from fuzzy_matcher import FuzzyMatcher
import pandas as pd

# Initialize the matcher
matcher = FuzzyMatcher(threshold=70)

# Load your data
source_df = pd.read_excel("source.xlsx")
target_df = pd.read_excel("target.xlsx")

# Perform matching
results = matcher.match_columns(
    source_df,
    target_df,
    source_column="SourceColumnName",
    target_column="TargetColumnName"
)

# Format results for export
results_df = matcher.format_results_for_export(results)

# Save to Excel
results_df.to_excel("matching_results.xlsx", index=False)
```

## Components

### 1. Synonym Handler (`synonym_handler.py`)
- Manages business terminology and abbreviations
- Integrates with WordNet for comprehensive synonym support
- Customizable synonym dictionary

### 2. Fuzzy Matcher (`fuzzy_matcher.py`)
- Core matching logic using fuzzy string matching
- Two-way matching algorithm
- Large dataset support using Dask
- Configurable matching threshold

### 3. Streamlit App (`streamlit_app.py`)
- User-friendly web interface
- File upload and SQL Server connection
- Interactive results visualization
- Excel report generation

## SQL Server SSO Authentication

The application uses Windows Authentication (SSO) for SQL Server connections. Ensure that:
1. You're running the application on a Windows machine
2. Your Windows account has necessary permissions on the SQL Server
3. SQL Server is configured to accept Windows Authentication

## Performance Considerations

- The application uses Dask for handling large datasets efficiently
- For very large SQL queries, consider adding appropriate indexes
- Adjust the Dask partition size in `fuzzy_matcher.py` if needed

## Results Format

The Excel output includes:

1. **Matching Results Sheet**
   - Type (Match/Mismatch)
   - Source Value
   - Target Value
   - DataItemID
   - Confidence Score
   - Matching Direction

2. **Summary Sheet**
   - Total Records Processed
   - Successful Matches
   - Source Mismatches
   - Target Mismatches
   - Average Confidence Score

## Customization

### Adding Custom Synonyms

Edit the `custom_synonyms` dictionary in `synonym_handler.py`:

```python
self.custom_synonyms = {
    'amt': ['amount'],
    'qty': ['quantity'],
    # Add your custom synonyms here
}
```

### Adjusting Matching Threshold

The default threshold is 70%. Adjust it when initializing the FuzzyMatcher:

```python
matcher = FuzzyMatcher(threshold=80)  # More strict matching
```

## Troubleshooting

1. **SQL Connection Issues**
   - Verify SQL Server name and database name
   - Ensure Windows Authentication is enabled
   - Check network connectivity

2. **Performance Issues**
   - Reduce dataset size or increase Dask partitions
   - Add appropriate indexes to SQL tables
   - Adjust matching threshold

3. **Memory Issues**
   - Increase available RAM
   - Reduce chunk size in Dask operations
   - Process data in batches

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
