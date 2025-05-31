import pandas as pd
import json
import os

def convert_excel_to_json():
    """
    Convert IREB practice exam Excel to JSON format
    Uses the exact filename from your folder
    """
    excel_file = "CorrectionAidForThePracticeExam_EN_2.0.1.xlsx"
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"File {excel_file} not found!")
        print("Make sure you're running this script in the same folder as the Excel file")
        return None
    
    try:
        # Read all sheets to see structure
        xl_file = pd.ExcelFile(excel_file)
        print(f"Sheet names: {xl_file.sheet_names}")
        
        # Read the first sheet
        df = pd.read_excel(excel_file, sheet_name=0)
        
        print("\n" + "="*50)
        print("EXCEL FILE STRUCTURE ANALYSIS")
        print("="*50)
        print(f"Shape: {df.shape} (rows, columns)")
        print(f"\nColumn names:")
        for i, col in enumerate(df.columns):
            print(f"  {i}: {col}")
        
        print(f"\nFirst 3 rows of data:")
        print(df.head(3).to_string())
        
        print(f"\nSample values from each column:")
        for col in df.columns[:6]:  # Show first 6 columns
            sample_val = str(df[col].dropna().iloc[0] if not df[col].dropna().empty else "No data")
            print(f"  {col}: {sample_val[:50]}...")
        
        # This will help us understand the structure before converting
        print("\n" + "="*50)
        print("NEXT STEP: Tell me which columns contain:")
        print("1. Question text")
        print("2. Answer options (A, B, C, D)")
        print("3. Correct answer")
        print("4. Any topic/learning objective info")
        print("="*50)
        
        return df
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

if __name__ == "__main__":
    convert_excel_to_json()
