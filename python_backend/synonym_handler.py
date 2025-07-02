import nltk
from nltk.corpus import wordnet
import re

class SynonymHandler:
    def __init__(self):
        # Download required NLTK data
        try:
            nltk.download('wordnet', quiet=True)
            nltk.download('averaged_perceptron_tagger', quiet=True)
        except:
            pass

        # Common business abbreviations and their expansions
        self.custom_synonyms = {
            'amt': ['amount'],
            'qty': ['quantity'],
            'id': ['identifier', 'identification'],
            'num': ['number'],
            'desc': ['description'],
            'acct': ['account'],
            'bal': ['balance'],
            'cust': ['customer'],
            'prod': ['product'],
            'ref': ['reference'],
            # Add more domain-specific synonyms as needed
        }

    def get_synonyms(self, word: str) -> set:
        """Get all possible synonyms for a word including custom business terms."""
        word = word.lower()
        synonyms = set()

        # Check custom business synonyms
        for term, expansions in self.custom_synonyms.items():
            if term == word:
                synonyms.update(expansions)
            elif word in expansions:
                synonyms.add(term)

        # Get WordNet synonyms
        for syn in wordnet.synsets(word):
            for lemma in syn.lemmas():
                synonyms.add(lemma.name().lower())

        # Add the original word
        synonyms.add(word)
        return synonyms

    def preprocess_column_name(self, column_name: str) -> str:
        """Preprocess column name for better matching."""
        # Convert to lowercase and remove special characters
        processed = re.sub(r'[^a-zA-Z0-9\s]', ' ', column_name.lower())
        # Remove extra spaces
        processed = ' '.join(processed.split())
        return processed

    def get_expanded_terms(self, column_name: str) -> set:
        """Get all possible variations of a column name including its parts."""
        processed_name = self.preprocess_column_name(column_name)
        terms = set()
        
        # Add the full name
        terms.add(processed_name)
        
        # Split and get synonyms for each part
        parts = processed_name.split()
        for part in parts:
            terms.update(self.get_synonyms(part))
            
        # Create combinations of adjacent terms
        for i in range(len(parts)-1):
            combined = f"{parts[i]}{parts[i+1]}"
            terms.add(combined)
            
        return terms
