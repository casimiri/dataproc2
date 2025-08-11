#!/usr/bin/env python3
"""
Test script to verify AI-powered species parsing functionality.
Run this with: OPENAI_API_KEY='your-key' python3 test_ai_species.py
"""

import os
import pandas as pd
from openai import OpenAI
from excelproc import split_latin_species_with_ai, split_latin_species

def test_ai_species_parsing():
    """Test the AI species parsing functionality"""
    
    # Check if API key is available
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        print("ERROR: OPENAI_API_KEY environment variable not found.")
        print("Please set it with: export OPENAI_API_KEY='your-api-key-here'")
        return False
    
    print("✓ OpenAI API key found")
    
    # Initialize OpenAI client
    try:
        client = OpenAI(api_key=api_key)
        print("✓ OpenAI client initialized successfully")
    except Exception as e:
        print(f"ERROR: Failed to initialize OpenAI client: {e}")
        return False
    
    # Load species reference
    try:
        species_df = pd.read_excel('Species.xlsx')
        species_list = species_df['Latin Name species'].tolist()
        print(f"✓ Loaded {len(species_list)} reference species")
    except Exception as e:
        print(f"ERROR: Failed to load Species.xlsx: {e}")
        return False
    
    print("\n=== Testing AI Species Parsing ===\n")
    
    # Test cases
    test_cases = [
        ("Pisum sativum", "Single species - should not split"),
        ("Pisum sativum Phaseolus vulgaris", "Two space-separated species"),
        ("Triticum aestivum Oryza sativa Glycine max", "Three space-separated species"),
        ("Hordeum vulgare, Rhanterium epapposum", "Comma-separated species"),
        ("Manihot esculenta Zea mays", "Two known species"),
    ]
    
    for i, (test_input, description) in enumerate(test_cases, 1):
        print(f"Test {i}: {description}")
        print(f"Input: '{test_input}'")
        
        try:
            # Test AI parsing
            ai_result = split_latin_species_with_ai(test_input, species_list, client)
            print(f"AI Result: {ai_result}")
            
            # Validate results
            validation = []
            for species in ai_result:
                if species in species_list:
                    validation.append(f"✓ {species}")
                else:
                    validation.append(f"? {species}")
            
            print(f"Validation: {', '.join(validation)}")
            print(f"Action: {'SPLIT into {len(ai_result)} rows' if len(ai_result) > 1 else 'NO SPLIT'}")
            
        except Exception as e:
            print(f"ERROR in AI parsing: {e}")
        
        print("-" * 60)
    
    print("\n=== Testing Full Pipeline ===")
    
    # Test the full split_latin_species function
    test_input = "Triticum aestivum Oryza sativa"
    print(f"Testing full pipeline with: '{test_input}'")
    
    try:
        result = split_latin_species(test_input, species_list, client)
        print(f"Full pipeline result: {result}")
        print(f"Success: AI-powered parsing {'WORKED' if len(result) > 1 else 'returned single item'}")
    except Exception as e:
        print(f"ERROR in full pipeline: {e}")
    
    return True

if __name__ == '__main__':
    print("=== OpenAI Species Parsing Test ===")
    test_ai_species_parsing()