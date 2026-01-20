#!/usr/bin/env python3
"""
Analyze KCSE performance by gender.
Note: Gender is inferred from names and may not be 100% accurate.
"""

import pandas as pd
import numpy as np
from collections import Counter

# Grade to points mapping (KCSE system)
grade_to_points = {
    'A': 12, 'A-': 11,
    'B+': 10, 'B': 9, 'B-': 8,
    'C+': 7, 'C': 6, 'C-': 5,
    'D+': 4, 'D': 3, 'D-': 2,
    'E': 1
}

def infer_gender_from_name(name):
    """
    Infer gender from name using common Kenyan naming patterns.
    This is a heuristic approach and may not be 100% accurate.
    """
    if pd.isna(name):
        return 'UNKNOWN'
    
    name_upper = str(name).upper()
    name_parts = name_upper.split()
    
    # Common male indicators (Kenyan context)
    male_indicators = [
        'JACKSON', 'JOSEPH', 'CHRISTOPHER', 'PETER', 'ANTONY', 'REGAN',
        'CLINTON', 'TYRUS', 'WILFRED', 'ALEX', 'BRITONY',
        'MOTERA', 'KERAWA', 'SINDA', 'AZERE', 'MWITA', 'BAGENI',
        'BUSH', 'BAREZY', 'RIOBA', 'MURIMI'
    ]
    
    # Common female indicators (Kenyan context)  
    female_indicators = [
        'MARY', 'ANN', 'JANE', 'ESTHER', 'GRACE', 'FAITH', 'HOPE',
        'SARAH', 'RUTH', 'RACHEL', 'REBECCA', 'SUSAN', 'CATHERINE',
        'DOROTHY', 'BEATRICE', 'AGNES', 'LUCY', 'ROSE', 'ELIZABETH'
    ]
    
    # Check first name
    first_name = name_parts[0] if name_parts else ""
    
    # Check against indicators
    if any(indicator in first_name for indicator in male_indicators):
        return 'MALE'
    elif any(indicator in first_name for indicator in female_indicators):
        return 'FEMALE'
    
    # Check middle names
    if len(name_parts) > 1:
        for part in name_parts[1:]:
            if any(indicator in part for indicator in male_indicators):
                return 'MALE'
            elif any(indicator in part for indicator in female_indicators):
                return 'FEMALE'
    
    # If uncertain, mark as UNKNOWN
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
    
    # Infer gender from names
    print("Inferring gender from names...")
    df['GENDER'] = df['NAME'].apply(infer_gender_from_name)
    
    # Count gender distribution
    gender_counts = df['GENDER'].value_counts()
    
    print("="*60)
    print("KCSE PERFORMANCE ANALYSIS BY GENDER - 2025")
    print("="*60)
    print()
    print("Gender Distribution:")
    print("-"*60)
    for gender, count in gender_counts.items():
        pct = (count / len(df)) * 100
        print(f"{gender:15s}: {count:4d} students ({pct:5.1f}%)")
    print(f"{'TOTAL':15s}: {len(df):4d} students")
    print()
    
    # Analyze by gender
    gender_analysis = {}
    
    for gender in ['MALE', 'FEMALE', 'UNKNOWN']:
        gender_df = df[df['GENDER'] == gender]
        
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
    print("="*60)
    
    for gender in ['MALE', 'FEMALE', 'UNKNOWN']:
        if gender not in gender_analysis:
            continue
        
        data = gender_analysis[gender]
        print()
        print(f"{gender} PERFORMANCE:")
        print("-"*60)
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
    print()
    print("="*60)
    print("GENDER COMPARISON:")
    print("="*60)
    
    if 'MALE' in gender_analysis and 'FEMALE' in gender_analysis:
        male_data = gender_analysis['MALE']
        female_data = gender_analysis['FEMALE']
        
        print()
        print(f"{'Metric':<25} {'MALE':<15} {'FEMALE':<15}")
        print("-"*60)
        print(f"{'Total Students':<25} {male_data['count']:<15} {female_data['count']:<15}")
        print(f"{'AB Count':<25} {male_data['ab_count']:<15} {female_data['ab_count']:<15}")
        print(f"{'AB %':<25} {(male_data['ab_count']/male_data['count']*100):.1f}%{'':<12} {(female_data['ab_count']/female_data['count']*100):.1f}%")
        
        if male_data['mean_points'] and female_data['mean_points']:
            print(f"{'Mean Points':<25} {male_data['mean_points']:.2f}{'':<12} {female_data['mean_points']:.2f}")
        
        print()
    
    print()
    print("="*60)
    print("NOTE: Gender classification is based on name patterns and may not be 100% accurate.")
    print("="*60)

if __name__ == "__main__":
    analyze_gender_performance()
