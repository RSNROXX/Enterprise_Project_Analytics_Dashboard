import pandas as pd
import os

# SET THIS TO YOUR ACTUAL FILE PATH
FILE_PATH = 'C:\Users\rishi\Downloads\Example Report _14 Jan 2026_05.14 pm.xlsx' # <--- CHANGE THIS to your actual file name

if not os.path.exists(FILE_PATH):
    print(f"❌ File not found: {FILE_PATH}")
else:
    xls = pd.ExcelFile(FILE_PATH)
    print(f"\n📂 File Found: {FILE_PATH}")
    print("="*40)
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        print(f"\n📄 SHEET: {sheet}")
        print("-" * 20)
        # Print clean, lowercase headers
        clean_headers = [str(c).strip().lower() for c in df.columns]
        print(", ".join(clean_headers))
        print("-" * 20)