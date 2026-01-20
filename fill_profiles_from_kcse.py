#!/usr/bin/env python3
"""
Script to fill Upload Student Profiles.xlsx with data from Students Upload KCSE Results file.
Fills: Index Number, Gender, and other matching fields.
"""

import pandas as pd
from pathlib import Path
import re
from itertools import permutations

def normalize_name(name):
    """Normalize name for matching (remove extra spaces, convert to uppercase)."""
    if pd.isna(name):
        return None
    name_str = str(name).strip().upper()
    # Remove extra spaces
    name_str = re.sub(r'\s+', ' ', name_str)
    return name_str

def find_name_column(df, possible_names):
    """Find the name column in a dataframe."""
    for col in df.columns:
        col_lower = str(col).lower()
        for possible in possible_names:
            if possible.lower() in col_lower:
                return col
    return None

def find_best_match(profile_name, name_to_data):
    """Try to find best matching name with fuzzy matching, including name interchanges."""
    if not profile_name:
        return None
    
    profile_parts = [p.strip() for p in profile_name.split() if p.strip()]
    
    # Exact match first
    if profile_name in name_to_data:
        return name_to_data[profile_name]
    
    if len(profile_parts) >= 2:
        # Try all permutations of name parts (for name interchanges)
        for perm in permutations(profile_parts):
            perm_name = ' '.join(perm)
            if perm_name in name_to_data:
                return name_to_data[perm_name]
        
        # Try matching with reversed order (most common interchange)
        if len(profile_parts) == 2:
            reversed_name = f"{profile_parts[1]} {profile_parts[0]}"
            if reversed_name in name_to_data:
                return name_to_data[reversed_name]
        elif len(profile_parts) == 3:
            # Try common 3-part name variations
            variations = [
                f"{profile_parts[2]} {profile_parts[0]} {profile_parts[1]}",  # LAST FIRST MIDDLE
                f"{profile_parts[2]} {profile_parts[1]} {profile_parts[0]}",  # LAST MIDDLE FIRST
                f"{profile_parts[0]} {profile_parts[2]} {profile_parts[1]}",  # FIRST LAST MIDDLE
            ]
            for var in variations:
                if var in name_to_data:
                    return name_to_data[var]
        
        # Try matching by last name + first name (swapped)
        last_first = f"{profile_parts[-1]} {profile_parts[0]}"
        if last_first in name_to_data:
            return name_to_data[last_first]
        
        # Try matching last name + first initial (swapped)
        if len(profile_parts[0]) > 0:
            last_first_initial = f"{profile_parts[-1]} {profile_parts[0][0]}"
            for kcse_name in name_to_data.keys():
                kcse_parts = kcse_name.split()
                if len(kcse_parts) >= 2:
                    if kcse_parts[0] == profile_parts[-1] and kcse_parts[1].startswith(profile_parts[0][0]):
                        return name_to_data[kcse_name]
        
        # Try matching by last name only (if unique)
        last_name = profile_parts[-1]
        matches = [name for name in name_to_data.keys() if name.endswith(last_name) or name.startswith(last_name)]
        if len(matches) == 1:
            return name_to_data[matches[0]]
    
    # Try partial matching (all parts must appear somewhere in KCSE name)
    for kcse_name in name_to_data.keys():
        kcse_parts = set(kcse_name.split())
        profile_parts_set = set(profile_parts)
        # Check if all significant parts (length > 2) appear in KCSE name
        significant_parts = [p for p in profile_parts if len(p) > 2]
        if significant_parts and all(any(p in kp for kp in kcse_parts) for p in significant_parts):
            # Additional check: at least 2 parts must match exactly
            exact_matches = len([p for p in profile_parts if p in kcse_parts])
            if exact_matches >= 2:
                return name_to_data[kcse_name]
    
    # Try fuzzy matching with spelling variations
    for kcse_name in name_to_data.keys():
        kcse_parts_list = kcse_name.split()
        if len(kcse_parts_list) == len(profile_parts):
            matches = 0
            for i, profile_part in enumerate(profile_parts):
                if i < len(kcse_parts_list):
                    kcse_part = kcse_parts_list[i]
                    # Exact match
                    if profile_part == kcse_part:
                        matches += 1
                    # Single character difference (common typos)
                    elif len(profile_part) == len(kcse_part) and len(profile_part) > 3:
                        diff_count = sum(1 for a, b in zip(profile_part, kcse_part) if a != b)
                        if diff_count <= 1:  # Only 1 character different
                            matches += 1
                    # Check if one contains the other (for abbreviations)
                    elif profile_part in kcse_part or kcse_part in profile_part:
                        matches += 1
            
            # If at least 2 parts match (including fuzzy matches)
            if matches >= 2:
                return name_to_data[kcse_name]
    
    # Try set-based matching with spelling tolerance
    profile_parts_set = set(profile_parts)
    for kcse_name in name_to_data.keys():
        kcse_parts_set = set(kcse_name.split())
        # Check for significant overlap
        common = profile_parts_set.intersection(kcse_parts_set)
        if len(common) >= 2:
            # Check if remaining parts are similar (spelling variations)
            remaining_profile = profile_parts_set - common
            remaining_kcse = kcse_parts_set - common
            if len(remaining_profile) <= 1 and len(remaining_kcse) <= 1:
                # Check if the remaining parts are similar
                if not remaining_profile or not remaining_kcse:
                    return name_to_data[kcse_name]
                # Check if remaining parts are similar (single char diff or substring)
                rp = list(remaining_profile)[0]
                rk = list(remaining_kcse)[0]
                if rp in rk or rk in rp or (len(rp) == len(rk) and sum(1 for a, b in zip(rp, rk) if a != b) <= 1):
                    return name_to_data[kcse_name]
    
    return None

def fill_profiles_from_kcse():
    """Fill Upload Student Profiles with data from KCSE Results file."""
    script_dir = Path(__file__).parent
    
    # File paths
    profiles_file = script_dir / "Upload Student Profiles.xlsx"
    kcse_file = script_dir / "Students Upload KCSE Results - Template.xlsx"
    
    # Check if updated version exists
    kcse_updated = script_dir / "Students Upload KCSE Results - Template_updated.xlsx"
    if kcse_updated.exists():
        kcse_file = kcse_updated
        print(f"Using updated KCSE file: {kcse_file.name}")
    
    if not profiles_file.exists():
        print(f"Error: Profiles file not found: {profiles_file}")
        return False
    
    if not kcse_file.exists():
        print(f"Error: KCSE file not found: {kcse_file}")
        return False
    
    # Read both Excel files
    print(f"Reading {profiles_file.name}...")
    profiles_df = pd.read_excel(profiles_file)
    print(f"  Rows: {len(profiles_df)}")
    
    print(f"\nReading {kcse_file.name}...")
    kcse_df = pd.read_excel(kcse_file)
    print(f"  Rows: {len(kcse_df)}")
    
    # Find name columns
    name_col_profiles = find_name_column(profiles_df, ['name', 'student name', 'full name', 'student'])
    name_col_kcse = find_name_column(kcse_df, ['name', 'student name', 'full name', 'student'])
    
    # Find index number columns
    index_col_profiles = find_name_column(profiles_df, ['index number', 'index no', 'index', 'index_number'])
    index_col_kcse = find_name_column(kcse_df, ['index number', 'index no', 'index', 'index_number'])
    
    # Find gender columns
    gender_col_profiles = find_name_column(profiles_df, ['gender', 'sex'])
    gender_col_kcse = find_name_column(kcse_df, ['gender', 'sex'])
    
    if not name_col_profiles:
        print("\nError: Could not find name column in Upload Student Profiles.xlsx")
        return False
    
    if not name_col_kcse:
        print("\nError: Could not find name column in KCSE Results file")
        return False
    
    if not index_col_kcse:
        print("\nError: Could not find Index Number column in KCSE Results file")
        return False
    
    print(f"\nUsing columns:")
    print(f"  Profiles - Name: {name_col_profiles}")
    print(f"  Profiles - Index: {index_col_profiles if index_col_profiles else '(will create)'}")
    print(f"  Profiles - Gender: {gender_col_profiles if gender_col_profiles else '(will create)'}")
    print(f"  KCSE - Name: {name_col_kcse}")
    print(f"  KCSE - Index: {index_col_kcse}")
    print(f"  KCSE - Gender: {gender_col_kcse if gender_col_kcse else '(not found)'}")
    
    # Create Index Number column in profiles if it doesn't exist
    if not index_col_profiles:
        index_col_profiles = 'Index Number'
        profiles_df[index_col_profiles] = None
    
    # Create Gender column in profiles if it doesn't exist
    if not gender_col_profiles:
        gender_col_profiles = 'GENDER'
        profiles_df[gender_col_profiles] = None
    
    # Convert Index Number column to string type to avoid dtype warnings
    profiles_df[index_col_profiles] = profiles_df[index_col_profiles].astype(str)
    profiles_df[index_col_profiles] = profiles_df[index_col_profiles].replace('nan', '')
    
    # Convert Gender column to string type to avoid dtype warnings
    if gender_col_profiles in profiles_df.columns:
        profiles_df[gender_col_profiles] = profiles_df[gender_col_profiles].astype(str)
        profiles_df[gender_col_profiles] = profiles_df[gender_col_profiles].replace('nan', '')
    
    # Create a mapping from normalized names to KCSE data
    name_to_data = {}
    for idx, row in kcse_df.iterrows():
        name = normalize_name(row[name_col_kcse])
        if name:
            index_num = row[index_col_kcse]
            index_str = None
            if pd.notna(index_num):
                index_str = str(index_num).strip()
                if index_str.endswith('.0'):
                    index_str = index_str[:-2]
            
            gender = None
            if gender_col_kcse and pd.notna(row[gender_col_kcse]):
                gender = str(row[gender_col_kcse]).strip().upper()
            
            name_to_data[name] = {
                'index': index_str,
                'gender': gender,
                'row': row  # Store entire row for potential future use
            }
    
    print(f"\nCreated name-to-data mapping with {len(name_to_data)} entries")
    
    # Match names and fill data
    matched_count = 0
    unmatched_names = []
    interchange_matches = []
    index_filled = 0
    gender_filled = 0
    
    for idx, row in profiles_df.iterrows():
        profile_name = normalize_name(row[name_col_profiles])
        current_index = str(row[index_col_profiles]).strip()
        current_gender = None
        if gender_col_profiles and pd.notna(row[gender_col_profiles]):
            current_gender = str(row[gender_col_profiles]).strip().upper()
        
        if not profile_name:
            continue
        
        # Try exact match first
        kcse_data = name_to_data.get(profile_name)
        match_type = "exact"
        
        if not kcse_data:
            # Try fuzzy matching
            kcse_data = find_best_match(profile_name, name_to_data)
            if kcse_data:
                # Check if this was an interchange match
                profile_parts = profile_name.split()
                if len(profile_parts) >= 2:
                    # Find the matched KCSE name
                    matched_kcse_name = None
                    for kcse_name, data in name_to_data.items():
                        if data == kcse_data:
                            matched_kcse_name = kcse_name
                            break
                    
                    if matched_kcse_name and matched_kcse_name != profile_name:
                        # Check if order is different
                        if set(profile_parts) == set(matched_kcse_name.split()):
                            match_type = "interchange"
                            interchange_matches.append((profile_name, matched_kcse_name))
        
        if kcse_data:
            # Fill Index Number
            if kcse_data['index'] and (not current_index or current_index == 'nan' or not current_index):
                profiles_df.at[idx, index_col_profiles] = kcse_data['index']
                index_filled += 1
            
            # Fill Gender
            if kcse_data['gender'] and (not current_gender or current_gender == 'NAN' or current_gender == ''):
                profiles_df.at[idx, gender_col_profiles] = kcse_data['gender']
                gender_filled += 1
            
            matched_count += 1
        else:
            unmatched_names.append(profile_name)
    
    # Print results
    # Count rows that already had index
    already_had_index = 0
    for idx, row in profiles_df.iterrows():
        current_index = str(row[index_col_profiles]).strip()
        if current_index and current_index != 'nan' and current_index:
            profile_name = normalize_name(row[name_col_profiles])
            if profile_name not in unmatched_names:
                already_had_index += 1
    
    print(f"\nResults:")
    print(f"  Total matched: {matched_count} rows")
    print(f"  Index Numbers filled: {index_filled}")
    print(f"  Gender filled: {gender_filled}")
    print(f"  Already had index: {already_had_index} rows")
    print(f"  Unmatched: {len(unmatched_names)} rows")
    
    if interchange_matches:
        print(f"\nMatches found via name interchange ({len(interchange_matches)}):")
        for profile_name, kcse_name in interchange_matches[:5]:
            print(f"  Profile: {profile_name}")
            print(f"  KCSE:    {kcse_name}")
        if len(interchange_matches) > 5:
            print(f"  ... and {len(interchange_matches) - 5} more")
    
    if unmatched_names:
        print(f"\nUnmatched names (first 10):")
        for name in unmatched_names[:10]:
            print(f"  - {name}")
        if len(unmatched_names) > 10:
            print(f"  ... and {len(unmatched_names) - 10} more")
    
    # Save the updated file
    output_file = profiles_file.parent / f"{profiles_file.stem}_filled_complete.xlsx"
    profiles_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"\nSaved updated file: {output_file.name}")
    
    return True

if __name__ == "__main__":
    fill_profiles_from_kcse()
