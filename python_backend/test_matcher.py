import pandas as pd
from .fuzzy_matcher import FuzzyMatcher
import tempfile
import os

def create_sample_data():
    """Create sample DataFrames for testing"""
    
    # Source data (similar to Table 1 in the example)
    source_data = {
        'ProjABS ID': ['ABS03', 'ABS10', 'ABS11', 'ABS14', 'ABS113'],
        'Attribute in ProjABS': [
            'Other Reverses',
            'Cash',
            'Cash and Cash equivalents',
            'Preffered Equity',
            'Other Income(Expense),Inclusive'
        ],
        'DataItemID': ['', '', '', '', '']
    }
    
    # Target data (similar to Table 2 in the example)
    target_data = {
        'DataItemID': ['1', '2', '3', '4', '5', '6'],
        'DataItemName': [
            'Cash(s)',
            'CashandCashequivalents',
            'Pref.Equity',
            'Property Expenses',
            '',
            ''
        ]
    }
    
    return pd.DataFrame(source_data), pd.DataFrame(target_data)

def test_matching():
    """Run a test matching process and display results"""
    
    print("Starting Fuzzy Column Matcher Test\n")
    
    # Create sample data
    source_df, target_df = create_sample_data()
    
    print("Source Data:")
    print(source_df)
    print("\nTarget Data:")
    print(target_df)
    
    # Initialize matcher
    matcher = FuzzyMatcher(threshold=70)
    
    print("\nRunning matching process...")
    
    # Perform matching
    results = matcher.match_columns(
        source_df,
        target_df,
        source_column='Attribute in ProjABS',
        target_column='DataItemName',
        id_column='DataItemID'
    )
    
    # Format results
    results_df = matcher.format_results_for_export(results)
    
    print("\nMatching Results:")
    print(results_df)
    
    # Save results to temporary file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        results_df.to_excel(tmp.name, index=False)
        print(f"\nResults saved to: {tmp.name}")
        
        # Display summary
        matches = len(results['matches'])
        source_mismatches = len(results['source_mismatches'])
        target_mismatches = len(results['target_mismatches'])
        
        print("\nSummary:")
        print(f"Total Matches: {matches}")
        print(f"Source Mismatches: {source_mismatches}")
        print(f"Target Mismatches: {target_mismatches}")
        
        # Calculate average confidence
        if matches > 0:
            avg_confidence = sum(m['confidence'] for m in results['matches']) / matches
            print(f"Average Match Confidence: {avg_confidence:.2f}%")

def test_synonym_matching():
    """Test the synonym matching capabilities"""
    
    print("\nTesting Synonym Matching\n")
    
    # Create test data with synonyms
    source_data = {
        'ID': ['1', '2', '3', '4'],
        'SourceColumn': [
            'Customer ID',
            'Amt Due',
            'Prod Description',
            'Acct Balance'
        ]
    }
    
    target_data = {
        'ID': ['1', '2', '3', '4'],
        'TargetColumn': [
            'Cust Identifier',
            'Amount Outstanding',
            'Product Desc',
            'Account Bal'
        ]
    }
    
    source_df = pd.DataFrame(source_data)
    target_df = pd.DataFrame(target_data)
    
    print("Source Data:")
    print(source_df)
    print("\nTarget Data:")
    print(target_df)
    
    # Initialize matcher
    matcher = FuzzyMatcher(threshold=60)  # Lower threshold for synonym matching
    
    print("\nRunning synonym matching...")
    
    # Perform matching
    results = matcher.match_columns(
        source_df,
        target_df,
        source_column='SourceColumn',
        target_column='TargetColumn'
    )
    
    # Format results
    results_df = matcher.format_results_for_export(results)
    
    print("\nSynonym Matching Results:")
    print(results_df)

if __name__ == "__main__":
    print("=== Fuzzy Column Matcher Tests ===\n")
    
    # Run main matching test
    test_matching()
    
    print("\n" + "="*30 + "\n")
    
    # Run synonym matching test
    test_synonym_matching()
    
    print("\nTests completed successfully!")
