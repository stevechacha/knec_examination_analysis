#!/usr/bin/env python3
"""
Analyze mean grade distribution from KCSE results.
"""

import pandas as pd

# Read the Excel file
df = pd.read_excel('Students Upload KCSE Results - Template_updated.xlsx')

# Get mean grades (remove NaN and convert to string)
mean_grades = df['MEAN_GRADE'].dropna().astype(str)

# Count distribution
grade_counts = mean_grades.value_counts()

# Calculate percentages
total = len(mean_grades)
percentages = (grade_counts / total * 100).round(1)

print('='*60)
print('KCSE MEAN GRADE DISTRIBUTION SUMMARY')
print('='*60)
print()
print(f'Total Students: {total}')
print()
print('Grade Distribution:')
print('-'*60)
print(f"{'Grade':<10} {'Count':<10} {'Percentage':<15}")
print('-'*60)

# Define grade order (A to E)
grade_order = ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']
for grade in grade_order:
    if grade in grade_counts.index:
        count = grade_counts[grade]
        pct = percentages[grade]
        bar = '█' * int(pct / 2)  # Visual bar
        print(f'{grade:<10} {count:<10} {pct:>6.1f}%  {bar}')

print('-'*60)

# Summary statistics
print()
print('Summary Statistics:')
print('-'*60)
print(f'Total Students: {total}')

# Group by grade bands
a_band = grade_counts[grade_counts.index.str.startswith('A')].sum() if any(grade_counts.index.str.startswith('A', na=False)) else 0
b_band = grade_counts[grade_counts.index.str.startswith('B')].sum() if any(grade_counts.index.str.startswith('B', na=False)) else 0
c_band = grade_counts[grade_counts.index.str.startswith('C')].sum() if any(grade_counts.index.str.startswith('C', na=False)) else 0
d_band = grade_counts[grade_counts.index.str.startswith('D')].sum() if any(grade_counts.index.str.startswith('D', na=False)) else 0
e_band = grade_counts.loc['E'] if 'E' in grade_counts.index else 0

print()
print('Grade Band Summary:')
print('-'*60)
print(f'A Band (A, A-):     {a_band:>3} students ({a_band/total*100:>5.1f}%)')
print(f'B Band (B+, B, B-): {b_band:>3} students ({b_band/total*100:>5.1f}%)')
print(f'C Band (C+, C, C-): {c_band:>3} students ({c_band/total*100:>5.1f}%)')
print(f'D Band (D+, D, D-): {d_band:>3} students ({d_band/total*100:>5.1f}%)')
print(f'E Grade:            {e_band:>3} students ({e_band/total*100:>5.1f}%)')

# Top performing grades
print()
print('Top Performing Grades:')
print('-'*60)
top_grades = grade_counts.head(3)
for grade, count in top_grades.items():
    pct = percentages[grade]
    print(f'{grade}: {count} students ({pct}%)')

print()
print('='*60)
