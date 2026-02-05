import PyPDF2
import pandas as pd
import re

# Get the PDF file path
pdf_file = r"c:\Users\HP\Desktop\data haraka\Rslt_Mvt_Enseignant_Prim2025.pdf"

# Initialize list to store all data
all_data = []

def clean_text(text):
    """Remove illegal characters while preserving Arabic text"""
    if not text:
        return text
    # Remove control characters but keep Arabic and common punctuation
    # Keep: Arabic characters, numbers, Latin letters, spaces, common punctuation
    cleaned = ''.join(char for char in text if ord(char) >= 32 or char in '\n\r\t')
    return cleaned.strip()

print("Opening PDF file with PyPDF2...")
try:
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        total_pages = len(pdf_reader.pages)
        print(f"Total pages: {total_pages}")
        
        # Extract text from all pages
        for page_num in range(total_pages):
            print(f"Processing page {page_num + 1}/{total_pages}...", end='\r')
            
            try:
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        cleaned_line = clean_text(line)
                        if cleaned_line:
                            all_data.append([cleaned_line])
                            
            except Exception as page_error:
                print(f"Warning: Could not fully process page {page_num + 1}: {str(page_error)[:50]}")
                continue
    
    print(" " * 100)  # Clear the progress line
    
    # Create DataFrame
    if all_data:
        df = pd.DataFrame(all_data, columns=['Content'])
        
        # Save to CSV first (UTF-8 encoding to preserve Arabic)
        csv_file = r"c:\Users\HP\Desktop\data haraka\output.csv"
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"✓ CSV file created: {csv_file}")
        
        # Save to Excel
        excel_file = r"c:\Users\HP\Desktop\data haraka\output.xlsx"
        df.to_excel(excel_file, index=False, engine='openpyxl')
        print(f"✓ Excel file created: {excel_file}")
        
        print(f"\nTotal rows extracted: {len(all_data)}")
        print("\nFirst 20 rows:")
        print(df.head(20).to_string())
    else:
        print("No data found in PDF")
        
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()

