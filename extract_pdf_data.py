"""
Script to extract data from report.pdf, save to file for review
"""
import pdfplumber

pdf_path = r"c:\Users\salu\Desktop\examtool 2\examtool\report.pdf"
output_path = r"c:\Users\salu\Desktop\examtool 2\examtool\pdf_extracted_data.txt"

with open(output_path, 'w', encoding='utf-8') as f:
    f.write("=" * 60 + "\n")
    f.write("Extracting data from report.pdf\n")
    f.write("=" * 60 + "\n")

    with pdfplumber.open(pdf_path) as pdf:
        f.write(f"\nTotal pages: {len(pdf.pages)}\n")
        
        all_data = []
        for i, page in enumerate(pdf.pages):
            f.write(f"\n{'='*60}\n")
            f.write(f"PAGE {i+1}\n")
            f.write("="*60 + "\n")
            
            # Extract tables
            tables = page.extract_tables()
            if tables:
                for j, table in enumerate(tables):
                    f.write(f"\n--- Table {j+1} ---\n")
                    for row in table:
                        f.write(str(row) + "\n")
                        all_data.append(row)
            else:
                # If no tables, extract text
                text = page.extract_text()
                if text:
                    f.write("\n--- Text Content ---\n")
                    f.write(text + "\n")

print(f"Data saved to {output_path}")
print("First 50 rows of extracted data:")
with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages[:2]):  # First 2 pages
        tables = page.extract_tables()
        if tables:
            for table in tables:
                for row in table[:25]:  # First 25 rows per page
                    print(row)
