import pdfplumber
import pandas as pd
from pathlib import Path

# Get the PDF file path
pdf_file = r"c:\Users\HP\Desktop\data haraka\Rslt_Mvt_Enseignant_Prim2025.pdf"

# Initialize list to store all data
all_data = []

print("Opening PDF file...")
try:
    with pdfplumber.open(pdf_file) as pdf:
        print(f"Total pages: {len(pdf.pages)}")
        
        # Extract text from all pages (without table extraction to avoid encoding issues)
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"Processing page {page_num}...", end='\r')
            
            try:
                # Extract raw text only - this preserves Arabic text better
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        line_stripped = line.strip()
                        if line_stripped:
                            all_data.append([line_stripped])
                    
            except Exception as page_error:
                print(f"Warning: Could not fully process page {page_num}: {str(page_error)[:50]}")
                continue
    
    print(" " * 80)  # Clear the progress line
    
    # Create DataFrame
    if all_data:
        df = pd.DataFrame(all_data, columns=['Content'])
        
        # Save to Excel
        excel_file = r"c:\Users\HP\Desktop\data haraka\output.xlsx"
        df.to_excel(excel_file, index=False, engine='openpyxl')
        print(f"✓ Excel file created: {excel_file}")
        
        # Save to CSV with UTF-8 encoding to preserve Arabic
        csv_file = r"c:\Users\HP\Desktop\data haraka\output.csv"
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"✓ CSV file created: {csv_file}")
        
        print(f"\nTotal rows extracted: {len(all_data)}")
        print("\nFirst 20 rows:")
        print(df.head(20).to_string())
    else:
        print("No data found in PDF")
        
except Exception as e:
    print(f"Error: {e}")
    print("Make sure you have the required libraries installed:")
    print("pip install pdfplumber pandas openpyxl")
