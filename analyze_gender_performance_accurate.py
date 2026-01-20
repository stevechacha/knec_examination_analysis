#!/usr/bin/env python3
"""
Analyze KCSE performance by gender using gender data from Excel file.
"""

import pandas as pd
import numpy as np

# Grade to points mapping (KCSE system)
grade_to_points = {
    'A': 12, 'A-': 11,
    'B+': 10, 'B': 9, 'B-': 8,
    'C+': 7, 'C': 6, 'C-': 5,
    'D+': 4, 'D': 3, 'D-': 2,
    'E': 1
}

def find_gender_column(df):
    """Find the gender column in the dataframe."""
    # Try common column names
    possible_names = ['GENDER', 'Gender', 'gender', 'SEX', 'Sex', 'sex', 
                      'GEND', 'Gend', 'gend', 'M/F', 'M_F', 'MALE/FEMALE']
    
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in possible_names:
            return col
        # Check if column name contains gender-related keywords
        col_lower = col_str.lower()
        if 'gender' in col_lower or 'sex' in col_lower or 'gend' in col_lower:
            return col
    
    return None

def normalize_gender(gender_value):
    """Normalize gender values to MALE/FEMALE/UNKNOWN."""
    if pd.isna(gender_value):
        return 'UNKNOWN'
    
    gender_str = str(gender_value).strip().upper()
    
    # Map various formats to standard values
    if gender_str in ['M', 'MALE', 'BOY', 'BOYS']:
        return 'MALE'
    elif gender_str in ['F', 'FEMALE', 'GIRL', 'GIRLS']:
        return 'FEMALE'
    else:
        return 'UNKNOWN'

def calculate_grade_distribution(grades):
    """Calculate grade distribution."""
    grades_clean = grades.dropna().astype(str)
    distribution = {}
    
    grade_order = ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']
    for grade in grade_order:
        distribution[grade] = len(grades_clean[grades_clean == grade])
    
    # Calculate AB (A and B bands)
    ab_count = distribution.get('A', 0) + distribution.get('A-', 0) + \
               distribution.get('B+', 0) + distribution.get('B', 0) + distribution.get('B-', 0)
    
    return distribution, ab_count

def calculate_mean_points(grades):
    """Calculate mean points from grades."""
    grades_clean = grades.dropna().astype(str)
    points_list = []
    
    for grade in grades_clean:
        if grade in grade_to_points:
            points_list.append(grade_to_points[grade])
    
    if points_list:
        return np.mean(points_list)
    return None

def analyze_gender_performance():
    """Analyze performance by gender."""
    
    # Load results
    df = pd.read_excel('Students Upload KCSE Results - Template_updated.xlsx')
    
    # Find gender column
    gender_col = find_gender_column(df)
    
    if gender_col is None:
        print("ERROR: Could not find gender column in the Excel file.")
        print(f"Available columns: {list(df.columns)}")
        print()
        print("Please ensure the Excel file has a gender column with one of these names:")
        print("  - GENDER, Gender, gender")
        print("  - SEX, Sex, sex")
        print("  - M/F, M_F, MALE/FEMALE")
        return
    
    print(f"Found gender column: '{gender_col}'")
    print()
    
    # Normalize gender values
    df['GENDER_NORMALIZED'] = df[gender_col].apply(normalize_gender)
    
    # Count gender distribution
    gender_counts = df['GENDER_NORMALIZED'].value_counts()
    
    print("="*70)
    print("KCSE PERFORMANCE ANALYSIS BY GENDER - 2025")
    print("="*70)
    print()
    print("Gender Distribution:")
    print("-"*70)
    for gender in ['MALE', 'FEMALE', 'UNKNOWN']:
        count = gender_counts.get(gender, 0)
        if count > 0:
            pct = (count / len(df)) * 100
            print(f"{gender:15s}: {count:4d} students ({pct:5.1f}%)")
    print(f"{'TOTAL':15s}: {len(df):4d} students")
    print()
    
    # Analyze by gender
    gender_analysis = {}
    
    for gender in ['MALE', 'FEMALE']:
        gender_df = df[df['GENDER_NORMALIZED'] == gender]
        
        if len(gender_df) == 0:
            continue
        
        mean_grades = gender_df['MEAN_GRADE'].dropna()
        
        # Calculate statistics
        grade_dist, ab_count = calculate_grade_distribution(mean_grades)
        mean_points = calculate_mean_points(mean_grades)
        
        gender_analysis[gender] = {
            'count': len(gender_df),
            'grade_dist': grade_dist,
            'ab_count': ab_count,
            'mean_points': mean_points,
            'mean_grades': mean_grades
        }
    
    # Print detailed analysis
    print("Performance by Gender:")
    print("="*70)
    
    for gender in ['MALE', 'FEMALE']:
        if gender not in gender_analysis:
            continue
        
        data = gender_analysis[gender]
        print()
        print(f"{gender} PERFORMANCE:")
        print("-"*70)
        print(f"Total Students: {data['count']}")
        print(f"AB Count: {data['ab_count']} ({(data['ab_count']/data['count']*100):.1f}%)")
        print(f"Mean Points: {data['mean_points']:.2f}" if data['mean_points'] else "Mean Points: N/A")
        
        print()
        print("Grade Distribution:")
        grade_order = ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']
        for grade in grade_order:
            count = data['grade_dist'].get(grade, 0)
            if count > 0:
                pct = (count / data['count']) * 100
                print(f"  {grade:3s}: {count:3d} students ({pct:5.1f}%)")
    
    # Comparison
    if 'MALE' in gender_analysis and 'FEMALE' in gender_analysis:
        print()
        print("="*70)
        print("GENDER COMPARISON:")
        print("="*70)
        
        male_data = gender_analysis['MALE']
        female_data = gender_analysis['FEMALE']
        
        print()
        print(f"{'Metric':<30} {'MALE':<20} {'FEMALE':<20}")
        print("-"*70)
        print(f"{'Total Students':<30} {male_data['count']:<20} {female_data['count']:<20}")
        print(f"{'AB Count':<30} {male_data['ab_count']:<20} {female_data['ab_count']:<20}")
        print(f"{'AB %':<30} {(male_data['ab_count']/male_data['count']*100):.1f}%{'':<17} {(female_data['ab_count']/female_data['count']*100):.1f}%")
        
        if male_data['mean_points'] and female_data['mean_points']:
            print(f"{'Mean Points':<30} {male_data['mean_points']:.2f}{'':<17} {female_data['mean_points']:.2f}")
        
        print()
        print("Grade Distribution Comparison:")
        print("-"*70)
        print(f"{'Grade':<10} {'MALE Count':<15} {'MALE %':<15} {'FEMALE Count':<15} {'FEMALE %':<15}")
        print("-"*70)
        grade_order = ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']
        for grade in grade_order:
            male_count = male_data['grade_dist'].get(grade, 0)
            female_count = female_data['grade_dist'].get(grade, 0)
            if male_count > 0 or female_count > 0:
                male_pct = (male_count / male_data['count']) * 100
                female_pct = (female_count / female_data['count']) * 100
                print(f"{grade:<10} {male_count:<15} {male_pct:>6.1f}%{'':<8} {female_count:<15} {female_pct:>6.1f}%")
    
    print()
    print("="*70)

if __name__ == "__main__":
    analyze_gender_performance()
