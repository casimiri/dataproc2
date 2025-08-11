import pandas as pd
import re
import sys
import argparse
import os
import json
from openai import OpenAI

def parse_varieties_with_openai(variety_text, client):
    """
    Use OpenAI GPT-3.5 Turbo to parse variety text and extract individual varieties.
    """
    if pd.isna(variety_text) or not isinstance(variety_text, str):
        return [variety_text]
    
    variety_text = variety_text.strip()
    if not variety_text:
        return [variety_text]
    
    prompt = f"""
Analyze the following plant variety text and extract individual varieties if multiple varieties are present.

Rules:
1. If the text contains only ONE variety (like "3. JIW.1"), return it as a single item
2. If the text contains MULTIPLE numbered varieties (like "1. SANEEN. 2. WQ 110" or "1. Lanet cocotype 2. Bisia ecohype 3. Kisia ecolipe"), extract each numbered variety separately
3. Return the result as a JSON array of strings
4. Keep the numbering (e.g., "1.", "2.", "3.") in the extracted varieties
5. Remove any trailing periods from variety names

Text to analyze: "{variety_text}"

Examples:
- Input: "3. JIW.1" → Output: ["3. JIW.1"]
- Input: "wheat(T.A) : 1. SANEEN. 2. WQ 110" → Output: ["1. SANEEN", "2. WQ 110"]
- Input: "Dolidos lablab Brachiaria nuzzizensis. 1. Lanet cocotype 2. Bisia ecohype 3. Kisia ecolipe" → Output: ["1. Lanet cocotype", "2. Bisia ecohype", "3. Kisia ecolipe"]

Return only the JSON array, no other text.
"""
    
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an expert at parsing plant variety names. Always return valid JSON arrays."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=500
        )
        
        result_text = response.choices[0].message.content.strip()
        
        # Try to parse the JSON response
        try:
            varieties = json.loads(result_text)
            if isinstance(varieties, list) and len(varieties) > 0:
                return varieties
        except json.JSONDecodeError:
            # If JSON parsing fails, try to extract array from text
            import re
            json_match = re.search(r'\[(.*?)\]', result_text, re.DOTALL)
            if json_match:
                try:
                    varieties = json.loads(f'[{json_match.group(1)}]')
                    if isinstance(varieties, list) and len(varieties) > 0:
                        return varieties
                except json.JSONDecodeError:
                    pass
    
    except Exception as e:
        print(f"OpenAI API error: {e}")
    
    # Fallback to original text if OpenAI fails
    return [variety_text]

def parse_varieties_with_regex(variety_text):
    """
    Fallback regex-based parsing for variety text.
    """
    if pd.isna(variety_text) or not isinstance(variety_text, str):
        return [variety_text]
    
    variety_text = variety_text.strip()
    
    # Function to extract numbered varieties from text
    def extract_numbered_varieties(text):
        # First try to find numbered varieties using regex
        pattern = r'(\d+\.\s*[^0-9]+?)(?=\s*\d+\.|$)'
        matches = re.findall(pattern, text)
        
        if len(matches) > 1:  # Multiple varieties found
            cleaned_varieties = []
            for match in matches:
                cleaned = match.strip().rstrip('.')
                if cleaned:
                    cleaned_varieties.append(cleaned)
            return cleaned_varieties
        
        # Fallback: try splitting approach for colon-separated cases
        parts = re.split(r'\.\s+(\d+\.)', text)
        
        if len(parts) > 1:  # Multiple varieties found
            varieties = []
            
            # First variety (before first split)
            first_variety = parts[0].strip()
            if first_variety:
                varieties.append(first_variety)
            
            # Remaining varieties
            for i in range(1, len(parts), 2):
                if i + 1 < len(parts):
                    variety_num = parts[i]  # e.g., "2."
                    variety_name = parts[i + 1].strip()  # e.g., "WQ 110"
                    if variety_name:
                        varieties.append(f"{variety_num} {variety_name}")
            
            return varieties if len(varieties) > 1 else None
        return None
    
    # Check if contains colon (indicating multiple varieties after colon)
    if ':' in variety_text:
        # Split by colon and process the part after colon
        parts = variety_text.split(':', 1)
        if len(parts) > 1:
            varieties_part = parts[1].strip()
            extracted = extract_numbered_varieties(varieties_part)
            if extracted:
                return extracted
    
    # Check for numbered varieties in the entire text (without colon)
    extracted = extract_numbered_varieties(variety_text)
    if extracted:
        return extracted
    
    # Single variety case (no numbered pattern found)
    return [variety_text]

def parse_varieties(variety_text, openai_client=None):
    """
    Parse variety text to extract individual varieties using OpenAI as primary method.
    Falls back to regex parsing if OpenAI is not available or fails.
    
    Examples:
    - "3. JIW.1" -> ["3. JIW.1"] (single variety)
    - "wheat(T.A) : 1. SANEEN. 2. WQ 110" -> ["1. SANEEN", "2. WQ 110"] (multiple varieties)
    - "Dolidos lablab Brachiaria nuzzizensis. 1. Lanet cocotype 2. Bisia ecohype 3. Kisia ecolipe" -> ["1. Lanet cocotype", "2. Bisia ecohype", "3. Kisia ecolipe"] (multiple varieties)
    
    Args:
        variety_text: The text to parse
        openai_client: OpenAI client instance (optional)
    
    Returns:
        List of varieties
    """
    if pd.isna(variety_text) or not isinstance(variety_text, str):
        return [variety_text]
    
    variety_text = variety_text.strip()
    if not variety_text:
        return [variety_text]
    
    # Try OpenAI first if client is available
    if openai_client:
        try:
            result = parse_varieties_with_openai(variety_text, openai_client)
            # Validate that we got a meaningful result
            if result and len(result) > 0 and result != [variety_text]:
                return result
        except Exception as e:
            print(f"OpenAI parsing failed, falling back to regex: {e}")
    
    # Fallback to regex parsing
    return parse_varieties_with_regex(variety_text)

def split_latin_species_with_ai(species_text, species_list, openai_client):
    """
    Use OpenAI to identify and extract individual species from text that may contain multiple species.
    
    Args:
        species_text: The text containing one or more species names
        species_list: List of valid species names from Species.xlsx
        openai_client: OpenAI client instance
    
    Returns:
        List of individual species names
    """
    if pd.isna(species_text) or not isinstance(species_text, str):
        return [species_text]
    
    species_text = species_text.strip()
    if not species_text:
        return [species_text]
    
    # Create a sample of species for the prompt to help AI understand the format
    species_sample = species_list[:20] if len(species_list) >= 20 else species_list
    
    prompt = f"""
Analyze the following text and extract individual Latin species names if multiple species are present.

Reference species list (sample): {', '.join(species_sample)}

Rules:
1. Look for individual Latin species names in the text (format: Genus species)
2. Species may be separated by commas, spaces, or other delimiters
3. Match extracted species against the reference list when possible
4. If the text contains only ONE species, return it as a single item
5. If the text contains MULTIPLE species, extract each one separately
6. Return the result as a JSON array of strings
7. Preserve the original species names as much as possible

Text to analyze: "{species_text}"

Examples:
- Input: "Pisum sativum" → Output: ["Pisum sativum"]
- Input: "Hordeum vulgare, Rhanterium epapposum" → Output: ["Hordeum vulgare", "Rhanterium epapposum"]
- Input: "Hordeum vulgare Rhanterium epapposum Calligonum polygonoides" → Output: ["Hordeum vulgare", "Rhanterium epapposum", "Calligonum polygonoides"]

Return only the JSON array, no other text.
"""
    
    try:
        response = openai_client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an expert botanist at parsing Latin species names. Always return valid JSON arrays."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=500
        )
        
        result_text = response.choices[0].message.content.strip()
        
        # Try to parse the JSON response
        try:
            species_names = json.loads(result_text)
            if isinstance(species_names, list) and len(species_names) > 0:
                return species_names
        except json.JSONDecodeError:
            # If JSON parsing fails, try to extract array from text
            import re
            json_match = re.search(r'\[(.*?)\]', result_text, re.DOTALL)
            if json_match:
                try:
                    species_names = json.loads(f'[{json_match.group(1)}]')
                    if isinstance(species_names, list) and len(species_names) > 0:
                        return species_names
                except json.JSONDecodeError:
                    pass
    
    except Exception as e:
        print(f"OpenAI API error for species parsing: {e}")
    
    # Fallback to original text if OpenAI fails
    return [species_text]

def split_latin_species_fallback(species_text, species_list):
    """
    Fallback method to split species using comma separation and simple heuristics.
    
    Args:
        species_text: The text containing one or more species names
        species_list: List of valid species names from Species.xlsx
    
    Returns:
        List of individual species names
    """
    if pd.isna(species_text) or not isinstance(species_text, str):
        return [species_text]
    
    species_text = species_text.strip()
    if not species_text:
        return [species_text]
    
    # Check for comma-separated species first
    if ',' in species_text:
        potential_species = [s.strip() for s in species_text.split(',')]
        if len(potential_species) > 1:
            return potential_species
    
    # Check for space-separated species using genus-species pattern
    # Look for pattern: "Genus species Genus species" etc.
    import re
    # Pattern to match Latin species names (Capitalized word followed by lowercase word)
    species_pattern = r'\b[A-Z][a-z]+\s+[a-z]+\b'
    matches = re.findall(species_pattern, species_text)
    
    if len(matches) > 1:
        # Validate matches against known species list
        valid_matches = []
        for match in matches:
            if match in species_list:
                valid_matches.append(match)
        
        if len(valid_matches) > 1:
            return valid_matches
        elif len(matches) > 1:
            # Even if not in species list, if we found multiple valid-looking species names, return them
            return matches
    
    # Return original text if no splitting occurred
    return [species_text]

def split_latin_species(species_text, species_list, openai_client=None):
    """
    Split Latin Name species text containing multiple species using AI or fallback methods.
    
    Args:
        species_text: The text containing one or more species names
        species_list: List of valid species names from Species.xlsx
        openai_client: OpenAI client instance (optional)
    
    Returns:
        List of individual species names
    """
    if pd.isna(species_text) or not isinstance(species_text, str):
        return [species_text]
    
    species_text = species_text.strip()
    if not species_text:
        return [species_text]
    
    # Try AI-based parsing first if available
    if openai_client:
        try:
            result = split_latin_species_with_ai(species_text, species_list, openai_client)
            # Validate that we got a meaningful result
            if result and len(result) > 0 and result != [species_text]:
                return result
        except Exception as e:
            print(f"AI species parsing failed, falling back to heuristics: {e}")
    
    # Fallback to heuristic-based parsing
    return split_latin_species_fallback(species_text, species_list)

def process_excel_file(input_file, output_file, species_file='Species.xlsx'):
    """
    Process Excel file to split rows with multiple Latin Name species using species reference file.
    """
    try:
        # Read the Species.xlsx file for reference
        try:
            species_df = pd.read_excel(species_file)
            species_list = species_df['Latin Name species'].tolist()
            print(f"Loaded {len(species_list)} reference species from {species_file}")
        except Exception as e:
            print(f"Warning: Could not load species reference file {species_file}: {e}")
            print("Proceeding without species validation.")
            species_list = []
        
        # Initialize OpenAI client
        openai_client = None
        api_key = os.getenv('OPENAI_API_KEY')
        
        if api_key:
            try:
                openai_client = OpenAI(api_key=api_key)
                print("OpenAI client initialized successfully. Using GPT-3.5 Turbo for variety parsing.")
            except Exception as e:
                print(f"Failed to initialize OpenAI client: {e}")
                print("Falling back to regex-based parsing.")
        else:
            print("OPENAI_API_KEY environment variable not found.")
            print("Set your OpenAI API key: export OPENAI_API_KEY='your-api-key-here'")
            print("Falling back to regex-based parsing.")
        
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        # Check available columns and determine which to process
        latin_species_column = 'Latin Name species'
        variety_column = 'Variety Name species'
        
        columns_to_process = []
        if latin_species_column in df.columns:
            columns_to_process.append(latin_species_column)
        if variety_column in df.columns:
            columns_to_process.append(variety_column)
            
        if not columns_to_process:
            print(f"Error: Neither '{latin_species_column}' nor '{variety_column}' columns found in the Excel file.")
            print(f"Available columns: {list(df.columns)}")
            return False
        
        print(f"Processing columns: {columns_to_process}")
        
        # Create new DataFrame to store processed data
        processed_rows = []
        total_rows = len(df)
        
        print(f"Processing {total_rows} rows...")
        
        for index, row in df.iterrows():
            if (index + 1) % 10 == 0 or (index + 1) == total_rows:
                print(f"Processing row {index + 1}/{total_rows}")
            
            # Process Latin Name species column
            if latin_species_column in df.columns:
                species_text = row[latin_species_column]
                species_list_for_row = split_latin_species(species_text, species_list, openai_client)
                
                if len(species_list_for_row) == 1:
                    # Single species or no splitting occurred
                    # Still process variety column if it exists
                    if variety_column in df.columns:
                        variety_text = row[variety_column]
                        varieties = parse_varieties(variety_text, openai_client)
                        
                        if len(varieties) == 1:
                            processed_rows.append(row)
                        else:
                            # Multiple varieties, create separate row for each
                            for variety in varieties:
                                new_row = row.copy()
                                new_row[variety_column] = variety
                                processed_rows.append(new_row)
                    else:
                        processed_rows.append(row)
                else:
                    # Multiple species found, create separate row for each species
                    for species in species_list_for_row:
                        # For each species, also check if there are multiple varieties
                        if variety_column in df.columns:
                            variety_text = row[variety_column]
                            varieties = parse_varieties(variety_text, openai_client)
                            
                            if len(varieties) == 1:
                                new_row = row.copy()
                                new_row[latin_species_column] = species
                                processed_rows.append(new_row)
                            else:
                                # Multiple varieties for each species
                                for variety in varieties:
                                    new_row = row.copy()
                                    new_row[latin_species_column] = species
                                    new_row[variety_column] = variety
                                    processed_rows.append(new_row)
                        else:
                            new_row = row.copy()
                            new_row[latin_species_column] = species
                            processed_rows.append(new_row)
            else:
                # Only process variety column if Latin Name species doesn't exist
                variety_text = row[variety_column]
                varieties = parse_varieties(variety_text, openai_client)
                
                if len(varieties) == 1:
                    processed_rows.append(row)
                else:
                    for variety in varieties:
                        new_row = row.copy()
                        new_row[variety_column] = variety
                        processed_rows.append(new_row)
        
        # Create new DataFrame from processed rows
        processed_df = pd.DataFrame(processed_rows)
        
        # Write to output file
        processed_df.to_excel(output_file, index=False)
        
        print(f"Successfully processed {len(df)} input rows into {len(processed_df)} output rows.")
        print(f"Output written to: {output_file}")
        
        return True
        
    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        return False
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Process Excel file to split multiple varieties into separate rows')
    parser.add_argument('input_file', help='Input Excel file path')
    parser.add_argument('output_file', help='Output Excel file path')
    
    args = parser.parse_args()
    
    success = process_excel_file(args.input_file, args.output_file)
    sys.exit(0 if success else 1)

if __name__ == '__main__':
    main()