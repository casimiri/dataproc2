#!/usr/bin/env python3
"""
Test script to demonstrate AI-powered species parsing functionality.
This script shows how the AI parsing would work if OpenAI API key was available.
"""

import pandas as pd
import json
from excelproc import split_latin_species_fallback

def mock_ai_species_parsing(species_text, species_list):
    """
    Mock AI parsing to demonstrate the functionality.
    This simulates what OpenAI would return for different test cases.
    """
    print(f"AI Parsing: '{species_text}'")
    
    # Mock responses for different patterns
    if species_text == "Pisum sativum Phaseolus vulgaris":
        return ["Pisum sativum", "Phaseolus vulgaris"]
    elif species_text == "Triticum aestivum Oryza sativa Glycine max":
        return ["Triticum aestivum", "Oryza sativa", "Glycine max"]
    elif "Dendranthema Chrysanthemum" in species_text:
        return ["Dendranthema", "Chrysanthemum spp."]
    else:
        # Use fallback for other cases
        return split_latin_species_fallback(species_text, species_list)

def main():
    # Load species reference
    species_df = pd.read_excel('Species.xlsx')
    species_list = species_df['Latin Name species'].tolist()
    
    print("=== AI-Powered Species Parsing Demo ===\n")
    
    # Test cases that demonstrate AI capabilities
    test_cases = [
        "Pisum sativum",  # Single species
        "Pisum sativum Phaseolus vulgaris",  # Space-separated known species
        "Triticum aestivum Oryza sativa Glycine max",  # Multiple known species
        "Dendranthema Chrysanthemum spp.",  # Mixed formats
        "Hordeum vulgare, Rhanterium epapposum, Calligonum polygonoides",  # Comma-separated
    ]
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"Test {i}: {test_case}")
        
        # Show fallback result
        fallback_result = split_latin_species_fallback(test_case, species_list)
        print(f"  Fallback: {fallback_result}")
        
        # Show mock AI result
        ai_result = mock_ai_species_parsing(test_case, species_list)
        print(f"  AI Mock:  {ai_result}")
        
        # Validate against species list
        valid_species = []
        for species in ai_result:
            if species in species_list:
                valid_species.append(f"✓ {species}")
            else:
                valid_species.append(f"? {species}")
        
        print(f"  Validation: {', '.join(valid_species)}")
        print(f"  Result: {'SPLIT' if len(ai_result) > 1 else 'NO SPLIT'}")
        print()

    print("=== Summary ===")
    print("✓ = Species found in reference list (Species.xlsx)")
    print("? = Species not found in reference list")
    print("\nWith OpenAI API key, the AI would:")
    print("- Parse complex species combinations more accurately")
    print("- Handle various delimiters and formats")
    print("- Match against the reference species list")
    print("- Preserve original species names when possible")

if __name__ == '__main__':
    main()