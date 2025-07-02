import pandas as pd
import dask.dataframe as dd
from fuzzywuzzy import fuzz
from typing import Dict, List, Tuple
import numpy as np
from synonym_handler import SynonymHandler

class FuzzyMatcher:
    def __init__(self, threshold: int = 70):
        self.threshold = threshold
        self.synonym_handler = SynonymHandler()

    def calculate_similarity(self, source: str, target: str) -> int:
        """
        Calculate similarity between two strings using fuzzy matching and synonyms.
        """
        # Get expanded terms for both strings
        source_terms = self.synonym_handler.get_expanded_terms(source)
        target_terms = self.synonym_handler.get_expanded_terms(target)
        
        # Calculate maximum similarity score among all term combinations
        max_score = 0
        for s_term in source_terms:
            for t_term in target_terms:
                score = fuzz.token_set_ratio(s_term, t_term)
                max_score = max(max_score, score)
        
        return max_score

    def match_columns(
        self,
        source_df: pd.DataFrame,
        target_df: pd.DataFrame,
        source_column: str,
        target_column: str,
        id_column: str = "DataItemID"
    ) -> Dict[str, List[Dict]]:
        """
        Perform two-way matching between source and target columns.
        Returns both matches and mismatches with confidence levels.
        """
        # Initialize results
        matches = []
        source_mismatches = []
        target_mismatches = []

        # Process source to target matches
        for _, source_row in source_df.iterrows():
            source_value = str(source_row[source_column]).strip()
            best_match = None
            best_score = -1
            
            # Find best match in target
            for _, target_row in target_df.iterrows():
                target_value = str(target_row[target_column]).strip()
                score = self.calculate_similarity(source_value, target_value)
                
                if score > best_score:
                    best_score = score
                    best_match = {
                        "value": target_value,
                        "id": target_row.get(id_column),
                        "confidence": score
                    }
            
            # Record match or mismatch
            if best_score >= self.threshold:
                matches.append({
                    "source_value": source_value,
                    "target_value": best_match["value"],
                    "data_item_id": best_match["id"],
                    "confidence": best_score,
                    "direction": "source_to_target"
                })
            else:
                source_mismatches.append({
                    "value": source_value,
                    "best_match": best_match["value"] if best_match else None,
                    "confidence": best_score,
                    "direction": "source_to_target"
                })

        # Process target to source matches (reverse direction)
        matched_target_values = {m["target_value"] for m in matches}
        
        for _, target_row in target_df.iterrows():
            target_value = str(target_row[target_column]).strip()
            
            # Skip if already matched
            if target_value in matched_target_values:
                continue
                
            best_match = None
            best_score = -1
            
            # Find best match in source
            for _, source_row in source_df.iterrows():
                source_value = str(source_row[source_column]).strip()
                score = self.calculate_similarity(target_value, source_value)
                
                if score > best_score:
                    best_score = score
                    best_match = {
                        "value": source_value,
                        "confidence": score
                    }
            
            # Record mismatch if no good match found
            if best_score < self.threshold:
                target_mismatches.append({
                    "value": target_value,
                    "id": target_row.get(id_column),
                    "best_match": best_match["value"] if best_match else None,
                    "confidence": best_score,
                    "direction": "target_to_source"
                })

        return {
            "matches": matches,
            "source_mismatches": source_mismatches,
            "target_mismatches": target_mismatches
        }

    def format_results_for_export(self, results: Dict[str, List[Dict]]) -> pd.DataFrame:
        """
        Format matching results into a pandas DataFrame suitable for export.
        """
        # Format matches
        match_records = [{
            "Type": "Match",
            "Source Value": m["source_value"],
            "Target Value": m["target_value"],
            "DataItemID": m["data_item_id"],
            "Confidence": f"{m['confidence']}%",
            "Direction": m["direction"]
        } for m in results["matches"]]
        
        # Format source mismatches
        source_mismatch_records = [{
            "Type": "Source Mismatch",
            "Source Value": m["value"],
            "Target Value": m["best_match"] if m["best_match"] else "No Match",
            "DataItemID": "N/A",
            "Confidence": f"{m['confidence']}%",
            "Direction": m["direction"]
        } for m in results["source_mismatches"]]
        
        # Format target mismatches
        target_mismatch_records = [{
            "Type": "Target Mismatch",
            "Source Value": m["best_match"] if m["best_match"] else "No Match",
            "Target Value": m["value"],
            "DataItemID": m["id"],
            "Confidence": f"{m['confidence']}%",
            "Direction": m["direction"]
        } for m in results["target_mismatches"]]
        
        # Combine all records
        all_records = match_records + source_mismatch_records + target_mismatch_records
        
        # Convert to DataFrame
        return pd.DataFrame(all_records)
