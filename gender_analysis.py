#!/usr/bin/env python3
import pandas as pd
import numpy as np

# Load data
df = pd.read_excel('Students Upload KCSE Results - Template_updated.xlsx')

# Grade to points
grade_to_points = {'A': 12, 'A-': 11, 'B+': 10, 'B': 9, 'B-': 8, 'C+': 7, 'C': 6, 'C-': 5, 'D+': 4, 'D': 3, 'D-': 2, 'E': 1}

# Simple gender inference
def get_gender(name):
    name_upper = str(name).upper()
    name_parts = name_upper.split()
    
    # Common male patterns (based on actual names in data)
    male_patts = ['JACKSON', 'JOSEPH', 'PETER', 'ANTONY', 'CHRISTOPHER', 'ALEX', 
                  'CLINTON', 'TYRUS', 'WILFRED', 'MOTERA', 'KERAWA', 'SINDA', 
                  'MWITA', 'BAGENI', 'BUSH', 'BAREZY', 'RIOBA', 'MURIMI', 
                  'AZERE', 'REGAN', 'BRITONY']
    
    for part in name_parts:
        if any(kw in part for kw in male_patts):
            return 'MALE'
    return 'UNKNOWN'

df['GENDER'] = df['NAME'].apply(get_gender)

# Analyze
print('='*60)
print('KCSE PERFORMANCE BY GENDER - 2025')
print('='*60)
print()

gender_counts = df['GENDER'].value_counts()
print('Gender Distribution:')
print('-'*60)
for gender, count in gender_counts.items():
    pct = (count / len(df)) * 100
    print(f'{gender:15s}: {count:4d} ({pct:5.1f}%)')
print(f'{"TOTAL":15s}: {len(df):4d}')
print()

# Performance by gender
print('Performance by Gender:')
print('='*60)

for gender in ['MALE', 'UNKNOWN']:
    gdf = df[df['GENDER'] == gender]
    if len(gdf) == 0:
        continue
    
    mg = gdf['MEAN_GRADE'].dropna().astype(str)
    
    # Points
    pts = [grade_to_points[g] for g in mg if g in grade_to_points]
    mean_pts = np.mean(pts) if pts else None
    
    # Grade dist
    grade_dist = mg.value_counts()
    
    # AB
    ab_grades = ['A', 'A-', 'B+', 'B', 'B-']
    ab_count = sum(grade_dist.get(g, 0) for g in ab_grades)
    
    print()
    print(f'{gender}:')
    print(f'  Total: {len(gdf)}')
    print(f'  AB: {ab_count} ({ab_count/len(gdf)*100:.1f}%)')
    if mean_pts:
        print(f'  Mean Points: {mean_pts:.2f}')
    print(f'  Grade Distribution:')
    for g in ['A', 'A-', 'B+', 'B', 'B-', 'C+', 'C', 'C-', 'D+', 'D', 'D-', 'E']:
        cnt = grade_dist.get(g, 0)
        if cnt > 0:
            print(f'    {g}: {cnt} ({cnt/len(gdf)*100:.1f}%)')

print()
print('='*60)
