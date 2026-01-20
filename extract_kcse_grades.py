#!/usr/bin/env python3
"""
Script to extract KCSE grades from screenshots and update Excel file.
"""

import os
import re
import glob
from pathlib import Path

try:
    import pandas as pd
    from PIL import Image
    import pytesseract
except ImportError as e:
    print(f"Error: Missing required package. {e}")
    print("Please install required packages:")
    print("  pip3 install pandas openpyxl pillow pytesseract")
    exit(1)

def extract_text_from_image(image_path):
    """Extract text from image using OCR."""
    try:
        image = Image.open(image_path)
        # Use --psm 6 for best results with structured text (like KCSE results)
        # This works best for subject-grade extraction
        text = pytesseract.image_to_string(image, config='--psm 6')
        return text
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return None

def extract_index_and_grades(text):
    """Extract index number and subject-grade pairs from OCR text."""
    # Common patterns for KCSE index numbers (typically 11 digits, e.g., 44748019001)
    # Try multiple patterns for index numbers
    index_patterns = [
        r'(\d{10,11}[/\-]\d{2,4})',  # Format: 44748019001/2019 or 4474801900/2019
        r'(\d{10,11})',  # Just 10-11 digit numbers
        r'(\d{6,10}[/\-]\d{2,4})',  # Format: 123456/2024 (fallback)
        r'(\d{6,10})',  # Just numbers (fallback)
    ]
    
    # Grade patterns (A, B+, B, B-, C+, C, C-, D+, D, D-, E)
    grade_pattern = r'\b([A-E][+-]?)\b'
    
    index_number = None
    grades = []  # Keep for backward compatibility
    subject_grades = {}  # Map subject names to grades
    
    # Find index number - try multiple patterns, prefer longer matches
    # Skip placeholder/example indices (1234567...)
    # Prioritize indices that appear near "KCSE", "Results", or student names
    best_match = None
    best_length = 0
    best_score = -1
    
    # First, try to find index numbers near "KCSE" or "Results" or followed by a dash and name
    # Pattern: index followed by dash and capital letters (student name)
    name_after_index_pattern = r'(\d{10,11})\s*-\s*[A-Z][A-Z\s]+'
    name_match = re.search(name_after_index_pattern, text)
    if name_match:
        potential_index = name_match.group(1)
        # Skip placeholder indices
        if not potential_index.startswith('123456'):
            best_match = potential_index
            best_length = len(potential_index)
            best_score = 100  # High priority
    
    # Also look for index numbers that match the Excel format (start with 4474801)
    # Only if we haven't already found a high-priority match (score 100)
    if best_score < 100:
        for pattern in index_patterns:
            index_matches = re.findall(pattern, text)
            for match in index_matches:
                match_str = str(match).split('/')[0]  # Remove year suffix for comparison
                # Skip placeholder indices (123456...)
                if match_str.startswith('123456'):
                    continue
                # Prioritize indices matching Excel format (4474801...)
                score = 10 if match_str.startswith('4474801') else 5
                match_length = len(match_str)
                # Prefer longer matches with higher scores
                if score > best_score or (score == best_score and match_length > best_length):
                    best_match = str(match)  # Keep original format if it has /year
                    best_length = match_length
                    best_score = score
    
    if best_match:
        index_number = best_match
        # If it doesn't have a year suffix but we found a year nearby, add it
        if '/' not in index_number and len(index_number) >= 6:
            year_match = re.search(r'(19|20)\d{2}', text)
            if year_match:
                index_number = f"{index_number}/{year_match.group()}"
    
    # Extract subject-grade pairs line by line for better accuracy
    # Pattern: number, code, subject name, grade (may have text after grade like "(PLAIN)")
    # Example: "1 101 ENGLISH C (PLAIN)" or "7 313 CHRISTIAN RELIGIOUS EDUCATION C (PLAIN)"
    # The challenge: "EDUCATION" starts with "E" which could be mistaken for grade "E"
    # Solution: Match grade that is followed by space+parentheses or end of meaningful text
    subject_matches = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Pattern: number, code, subject words, then grade followed by optional (PLAIN) etc
        # Use a pattern that ensures grade is followed by parentheses or end of line
        # Match: digits, digits, then subject (words), then grade, then optional (text)
        # The grade must be followed by whitespace and parentheses OR be at end of meaningful content
        pattern = r'^(\d+)\s+(\d+)\s+(.+?)\s+([A-E][+-]?|X)(?:\s*\([^)]+\))?\s*$'
        match = re.match(pattern, line, re.IGNORECASE)
        if match:
            subject_name = ' '.join(match.group(3).split()).strip()
            grade = match.group(4).strip()
            
            # Validate: grade should be a valid grade
            if grade in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E', 'X']:
                # Accept all valid grades - the line pattern already ensures grade comes after subject
                # For "AGRICULTURE E", the pattern correctly separates subject and grade
                subject_matches.append((subject_name, grade))
    
    # If line-by-line didn't work, try the original pattern approach
    if len(subject_matches) < 3:
        subject_patterns = [
            r'(\d+)\s+(\d+)\s+([A-Z][A-Z\s]+?)\s+([A-E][+-]?|X)(?:\s*\([^)]+\))?',
        ]
        for pattern in subject_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches and len(matches) >= 3:
                cleaned_matches = []
                for m in matches:
                    subject_name = ' '.join(m[2].split()).strip()
                    grade = m[3].strip()
                    if grade in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E', 'X']:
                        if not subject_name.upper().endswith(grade[0]):
                            cleaned_matches.append((subject_name, grade))
                if cleaned_matches:
                    subject_matches = cleaned_matches
                    break
    
    # Subject name mappings (common variations)
    subject_map = {
        'ENGLISH': 'ENG',
        'KISWAHILI': 'KIS',
        'MATHEMATICS': 'MAT',
        'BIOLOGY': 'BIO',
        'PHYSICS': 'PHY',
        'CHEMISTRY': 'CHE',
        'GEOGRAPHY': 'GEO',
        'HISTORY': 'HIS',
        'HISTORY AND GOVERNMENT': 'HIS',
        'CHRISTIAN RELIGIOUS EDUCATION': 'CRE',
        'CHRISTIAN RELIGIOUS': 'CRE',
        'AGRICULTURE': 'AGR',
        'BUSINESS STUDIES': 'BST',
        'BUSINESS': 'BST',
        'COMPUTER STUDIES': 'COM',
        'COMPUTER': 'COM',
    }
    
    for subject_name, grade in subject_matches:
        subject_name = subject_name.strip().upper()
        # Map subject name to column name
        mapped = False
        for key, col_name in subject_map.items():
            if key in subject_name:
                if grade != 'X' and grade.upper() != 'X':  # Skip X (not available)
                    subject_grades[col_name] = grade
                    grades.append(grade)  # Keep for backward compatibility
                mapped = True
                break
        # Debug: if not mapped, print it
        if not mapped and subject_name:
            pass  # Subject not in our map - that's okay
    
    # Fallback: if no subject-grade pairs found, extract all grades
    if not subject_grades:
        grade_matches = re.findall(grade_pattern, text)
        if grade_matches:
            valid_grades = ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']
            grades = [g for g in grade_matches if g in valid_grades]
    
    # Extract mean grade from text (pattern: "Mean Grade: B" or "Mean Grade:B" or "Mean Grade B")
    mean_grade = None
    mean_grade_patterns = [
        r'Mean\s+Grade[:\s]+([A-E][+-]?)\s*\(',  # Mean Grade: B (PLAIN)
        r'Mean\s+Grade[:\s]+([A-E][+-]?)\s',  # Mean Grade: B 
        r'Mean\s+Grade[:\s]+([A-E][+-]?)$',  # Mean Grade: B (end of line)
        r'Mean[:\s]+([A-E][+-]?)\s',  # Mean: B (simpler pattern)
    ]
    
    for pattern in mean_grade_patterns:
        mean_match = re.search(pattern, text, re.IGNORECASE)
        if mean_match:
            mean_grade = mean_match.group(1).strip()
            # Validate it's a valid grade
            if mean_grade in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']:
                break
    
    return index_number, grades, subject_grades, mean_grade

def process_all_screenshots(screenshot_dir):
    """Process all screenshot files and extract data."""
    screenshot_files = sorted(glob.glob(os.path.join(screenshot_dir, "Screenshot*.png")))
    results = {}  # Will store by index number
    all_extractions = []  # Store all extractions for analysis
    
    print(f"Found {len(screenshot_files)} screenshot files")
    
    for i, screenshot_path in enumerate(screenshot_files, 1):
        print(f"Processing {i}/{len(screenshot_files)}: {os.path.basename(screenshot_path)}")
        text = extract_text_from_image(screenshot_path)
        
        if text:
            index_number, grades, subject_grades, mean_grade = extract_index_and_grades(text)
            
            # Store all extractions
            all_extractions.append({
                'file': os.path.basename(screenshot_path),
                'index': index_number,
                'grades': grades,
                'subject_grades': subject_grades,
                'mean_grade': mean_grade,
                'text_preview': text[:300] if text else ''
            })
            
            if index_number:
                # Skip placeholder index numbers
                if '12345678' in index_number or '1234567' in index_number:
                    print(f"  Skipping placeholder index: {index_number}")
                    # Try to find actual index in text
                    # Look for patterns like 44748017xxx (matching Excel format)
                    better_match = re.search(r'4474801[789]\d{3}', text)
                    if better_match:
                        index_number = better_match.group()
                        print(f"  Found better match: {index_number}")
                
                if index_number and '12345678' not in index_number:
                    # Store by normalized index (without /2019)
                    normalized_idx = normalize_index_number(index_number)
                    if normalized_idx:
                        results[normalized_idx] = {
                            'grades': grades,
                            'subject_grades': subject_grades,  # Store subject-grade mapping
                            'mean_grade': mean_grade,  # Store mean grade
                            'source_file': os.path.basename(screenshot_path),
                            'original_index': index_number
                        }
                        subjects_str = ', '.join([f"{k}:{v}" for k, v in subject_grades.items()])
                        mean_str = f", Mean: {mean_grade}" if mean_grade else ""
                        print(f"  Found index: {index_number} -> {normalized_idx}, Subjects: {subjects_str}{mean_str}")
            else:
                print(f"  Could not extract index number from {os.path.basename(screenshot_path)}")
                # Try harder - look for any 11-digit number starting with 4474801[789] (matches Excel format)
                hard_match = re.search(r'4474801[789]\d{3}', text)
                if hard_match:
                    index_number = hard_match.group()
                    normalized_idx = normalize_index_number(index_number)
                    if normalized_idx:
                        results[normalized_idx] = {
                            'grades': grades,
                            'subject_grades': subject_grades,
                            'mean_grade': mean_grade,
                            'source_file': os.path.basename(screenshot_path),
                            'original_index': index_number
                        }
                        print(f"  Found index via hard search: {index_number}, Grades: {grades}, Mean: {mean_grade}")
                else:
                    print(f"  OCR text preview: {text[:300]}...")
        else:
            print(f"  Failed to extract text from {os.path.basename(screenshot_path)}")
    
    print(f"\nTotal unique index numbers extracted: {len(results)}")
    return results

def normalize_index_number(index_num):
    """Normalize index number format for matching."""
    if index_num is None or (hasattr(pd, 'isna') and pd.isna(index_num)):
        return None
    
    # Convert to string and handle float format (e.g., 44748019001.0 -> 44748019001)
    index_str = str(index_num).strip()
    
    # Remove .0 suffix if it's a float representation
    if index_str.endswith('.0'):
        index_str = index_str[:-2]
    
    # Remove any "/YYYY" year suffix (e.g., "/2019" -> "")
    index_str = re.sub(r'/\d{4}$', '', index_str)
    
    # Remove any extra whitespace or formatting
    index_str = re.sub(r'\s+', '', index_str)
    
    return index_str

def update_excel_file(excel_path, extracted_data):
    """Update Excel file with extracted grades."""
    try:
        # Read the Excel file
        df = pd.read_excel(excel_path)
        print(f"\nExcel file loaded. Columns: {df.columns.tolist()}")
        print(f"Number of rows: {len(df)}")
        
        # Find the index number column (common names: Index Number, Index No, Index, INDEX_NO, etc.)
        index_col = None
        for col in df.columns:
            col_lower = str(col).lower()
            if 'index' in col_lower and ('number' in col_lower or 'no' in col_lower or len([w for w in ['index', 'number', 'no'] if w in col_lower]) >= 1):
                index_col = col
                break
        
        if index_col is None:
            print("Could not find index number column. Available columns:")
            for i, col in enumerate(df.columns, 1):
                print(f"  {i}. {col}")
            print("\nPlease check the Excel file structure.")
            return False
        
        print(f"Using index column: {index_col}")
        
        # Create a mapping of normalized index numbers to row indices
        index_to_row = {}
        for idx, row in df.iterrows():
            index_num = normalize_index_number(row[index_col])
            if index_num:
                index_to_row[index_num] = idx
                # Also create partial matches (for OCR errors - e.g., if Excel has 44748019001 and OCR has 4474801900)
                # Try matching last 9-10 digits
                if len(index_num) >= 9:
                    partial_key = index_num[-9:]  # Last 9 digits
                    if partial_key not in index_to_row:
                        index_to_row[partial_key] = idx
                    partial_key = index_num[-10:]  # Last 10 digits
                    if partial_key not in index_to_row:
                        index_to_row[partial_key] = idx
        
        # Find subject columns (order chosen to match the typical KCSE results layout)
        # Priority order matches the screenshot: ENG, KIS, MAT, BIO, CHE, GEO, CRE, AGR, BST ...
        subject_columns = [
            'ENG', 'KIS', 'MAT', 'BIO', 'CHE', 'GEO', 'CRE', 'AGR', 'BST',
            # Additional subjects that may appear
            'HIS', 'PHY', 'COM', 'IRE', 'HRE', 'KSL', 'FAS', 'LIT', 'GSC', 'AD', 'HSC'
        ]
        
        available_subject_cols = [col for col in subject_columns if col in df.columns]
        
        print(f"Found {len(available_subject_cols)} subject columns: {available_subject_cols[:10]}...")
        
        # Normalize extracted data keys for matching
        normalized_extracted = {}
        for key, value in extracted_data.items():
            normalized_key = normalize_index_number(key)
            if normalized_key:
                normalized_extracted[normalized_key] = value
        
        # Update grades
        updated_count = 0
        for normalized_index, data in normalized_extracted.items():
            # Try exact match first
            row_idx = None
            if normalized_index in index_to_row:
                row_idx = index_to_row[normalized_index]
            else:
                # Try partial match (last 9-10 digits)
                if len(normalized_index) >= 9:
                    partial_key = normalized_index[-9:]
                    if partial_key in index_to_row:
                        row_idx = index_to_row[partial_key]
                    elif len(normalized_index) >= 10:
                        partial_key = normalized_index[-10:]
                        if partial_key in index_to_row:
                            row_idx = index_to_row[partial_key]
            
            if row_idx is not None:
                # Use subject-grade mapping if available, otherwise fall back to sequential
                subject_grades = data.get('subject_grades', {})
                grades = data.get('grades', [])
                mean_grade = data.get('mean_grade', None)
                
                if subject_grades:
                    # Map each subject to its specific column
                    updated_subjects = []
                    for subject_col, grade in subject_grades.items():
                        if subject_col in df.columns:
                            # Convert column to string type if needed
                            if df[subject_col].dtype != 'object':
                                df[subject_col] = df[subject_col].astype(str)
                            df.at[row_idx, subject_col] = str(grade)
                            updated_subjects.append(f"{subject_col}:{grade}")
                    
                    # Update mean_grade if it exists (use extracted mean_grade, not first grade)
                    mean_grade_col = None
                    if 'MEAN_GRADE' in df.columns:
                        mean_grade_col = 'MEAN_GRADE'
                    elif 'mean_grade' in df.columns:
                        mean_grade_col = 'mean_grade'
                    
                    if mean_grade_col and mean_grade:
                        if df[mean_grade_col].dtype != 'object':
                            df[mean_grade_col] = df[mean_grade_col].astype(str)
                        df.at[row_idx, mean_grade_col] = str(mean_grade)
                    
                    subjects_str = ', '.join(updated_subjects)
                    excel_index = df.at[row_idx, index_col]
                    print(f"Updated row {row_idx+1} (Excel index: {excel_index}) with {len(updated_subjects)} subjects: {subjects_str}")
                    updated_count += 1
                elif grades and available_subject_cols:
                    # Fallback: sequential mapping if subject-grade pairs not found
                    num_grades = len(grades)
                    num_cols = len(available_subject_cols)
                    
                    # Update subject columns with grades (convert to string to avoid dtype warnings)
                    for i, grade in enumerate(grades):
                        if i < num_cols:
                            col_name = available_subject_cols[i]
                            # Convert column to string type if needed
                            if df[col_name].dtype != 'object':
                                df[col_name] = df[col_name].astype(str)
                            df.at[row_idx, col_name] = str(grade)
                    
                    # Update mean_grade if it exists (use extracted mean_grade, not first grade)
                    mean_grade_col = None
                    if 'MEAN_GRADE' in df.columns:
                        mean_grade_col = 'MEAN_GRADE'
                    elif 'mean_grade' in df.columns:
                        mean_grade_col = 'mean_grade'
                    
                    if mean_grade_col and mean_grade:
                        if df[mean_grade_col].dtype != 'object':
                            df[mean_grade_col] = df[mean_grade_col].astype(str)
                        df.at[row_idx, mean_grade_col] = str(mean_grade)
                    
                    grades_str = ', '.join(grades) if grades else ''
                    excel_index = df.at[row_idx, index_col]
                    print(f"Updated row {row_idx+1} (Excel index: {excel_index}) with {len(grades)} grades (sequential): {grades_str}")
                    updated_count += 1
            else:
                print(f"Warning: Index {normalized_index} not found in Excel file")
        
        # Save the updated Excel file
        output_path = excel_path.replace('.xlsx', '_updated.xlsx')
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\nUpdated Excel file saved as: {output_path}")
        print(f"Successfully updated {updated_count} records")
        
        return True
    except Exception as e:
        print(f"Error updating Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    # Get the script directory
    script_dir = Path(__file__).parent
    excel_file = script_dir / "Students Upload KCSE Results - Template.xlsx"
    
    if not excel_file.exists():
        print(f"Error: Excel file not found: {excel_file}")
        return
    
    # Process all screenshots
    print("Extracting data from screenshots...")
    extracted_data = process_all_screenshots(str(script_dir))
    
    print(f"\nExtracted data from {len(extracted_data)} screenshots:")
    for index_num, data in extracted_data.items():
        print(f"  {index_num}: {data['grades']} (from {data['source_file']})")
    
    # Update Excel file
    print("\nUpdating Excel file...")
    update_excel_file(str(excel_file), extracted_data)

if __name__ == "__main__":
    main()
