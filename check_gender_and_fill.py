#!/usr/bin/env python3
"""Check for gender column and fill template."""

import pandas as pd
import sys

# Check the Excel file
df = pd.read_excel('Students Upload KCSE Results - Template_updated.xlsx')

print("="*60)
print("CHECKING GENDER COLUMN")
print("="*60)
print(f"\nAll columns in file:")
for i, col in enumerate(df.columns, 1):
    print(f"  {i}. {col}")
    
print(f"\nFirst 5 rows:")
print(df.head(5).to_string())

# Check for gender column
gender_col = None
for col in df.columns:
    col_lower = str(col).lower()
    if 'gender' in col_lower or 'sex' in col_lower or 'gend' in col_lower:
        gender_col = col
        print(f"\n✓ Found gender column: '{col}'")
        print(f"\nGender value counts:")
        print(df[col].value_counts())
        break

if not gender_col:
    print("\n✗ WARNING: No gender column found!")
    print("Please ensure your Excel file has a gender column.")
    sys.exit(1)

print("\n" + "="*60)
print("Running fill_analysis_template.py...")
print("="*60 + "\n")
