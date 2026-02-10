"""
Scharine Quote Scraper
Extracts item part numbers, descriptions, quantities, and prices from quote PDFs
"""

import os
import re
import pandas as pd
import PyPDF2
from pathlib import Path


def extract_quote_data(pdf_path):
    """
    Extract item information from a Scharine quote PDF
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        list of dicts containing item data
    """
    items = []
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            
            # Extract text from all pages
            for page in pdf_reader.pages:
                text += page.extract_text()
            
            # Split into lines for processing
            lines = text.split('\n')
            
            # Find the parts quote section
            in_parts_section = False
            current_item = {}
            
            for i, line in enumerate(lines):
                # Start of parts section
                if 'PARTS QUOTE' in line:
                    in_parts_section = True
                    continue
                
                # End of parts section (when we hit terms/totals)
                if in_parts_section and ('LEAD TIME' in line or 'Net Order:' in line):
                    break
                
                if in_parts_section:
                    # Pattern to match part numbers (e.g., F7300K-ZNC, F518525ASM-C)
                    part_pattern = r'^([A-Z0-9]+-[A-Z0-9-]+)\s*$'
                    
                    # Check if this line is a part number
                    part_match = re.match(part_pattern, line.strip())
                    
                    if part_match:
                        # Save previous item if exists
                        if current_item and 'part_number' in current_item:
                            items.append(current_item)
                        
                        # Start new item
                        current_item = {'part_number': part_match.group(1)}
                    
                    # Look for "Whse:" which precedes description and quantity
                    elif 'Whse:' in line:
                        # Extract warehouse and description
                        whse_pattern = r'Whse:\s+(\d+)\s+(.+)'
                        whse_match = re.search(whse_pattern, line)
                        if whse_match:
                            description = whse_match.group(2).strip()
                            current_item['description'] = description
                    
                    # Look for EACH quantity price pattern
                    elif 'EACH' in line:
                        # Pattern: EACH  quantity  shipped  back_order  price  amount
                        parts = line.split()
                        if len(parts) >= 3:
                            try:
                                qty_idx = parts.index('EACH') + 1
                                if qty_idx < len(parts):
                                    # Quantity is right after EACH
                                    current_item['quantity'] = float(parts[qty_idx])
                                    
                                    # Price is typically near the end
                                    # Look for the second-to-last or third-to-last number
                                    for j in range(len(parts) - 1, -1, -1):
                                        try:
                                            price = float(parts[j].replace(',', ''))
                                            # Check if this looks like a unit price (not too large)
                                            if price < 10000:
                                                current_item['price'] = price
                                                break
                                        except ValueError:
                                            continue
                            except (ValueError, IndexError):
                                pass
            
            # Add last item
            if current_item and 'part_number' in current_item:
                items.append(current_item)
    
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
    
    return items


def process_quotes_folder(folder_path):
    """
    Process all PDF files in the quotes folder
    
    Args:
        folder_path: Path to folder containing quote PDFs
        
    Returns:
        DataFrame with all extracted data
    """
    folder = Path(folder_path)
    all_items = []
    
    # Process each PDF file
    for pdf_file in folder.glob('*.pdf'):
        print(f"Processing: {pdf_file.name}")
        items = extract_quote_data(pdf_file)
        
        # Add source file to each item
        for item in items:
            item['source_file'] = pdf_file.name
        
        all_items.extend(items)
    
    # Create DataFrame
    df = pd.DataFrame(all_items)
    
    # Reorder columns
    if not df.empty:
        column_order = ['source_file', 'part_number', 'description', 'quantity', 'price']
        df = df[[col for col in column_order if col in df.columns]]
    
    return df


def main():
    # Set your folder path here
    folder_path = r"C:\Users\bromayo\OneDrive - Tesla\Quotes\Scharine Quotes"
    
    print(f"Scanning folder: {folder_path}")
    print("-" * 80)
    
    # Process all quotes
    df = process_quotes_folder(folder_path)
    
    # Display results
    if not df.empty:
        print(f"\nExtracted {len(df)} items from quotes")
        print("\nPreview:")
        print(df.to_string(index=False))
        
        # Save to Excel
        output_file = os.path.join(folder_path, 'extracted_quotes.xlsx')
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\nData saved to: {output_file}")
        
        # Also save as CSV
        csv_file = os.path.join(folder_path, 'extracted_quotes.csv')
        df.to_csv(csv_file, index=False)
        print(f"Data saved to: {csv_file}")
    else:
        print("No items extracted. Please check the PDF format.")


if __name__ == "__main__":
    main()
