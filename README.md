# Scharine Quote Scraper

This Python script extracts item information from Scharine quote PDFs, specifically:
- Part Number (e.g., F7300K-ZNC)
- Item Description (e.g., REAR HITCH, CLR ZINC)
- Quantity Ordered
- Unit Price

## Installation

1. Make sure you have Python 3.7+ installed
2. Open Command Prompt and navigate to your script folder
3. Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

1. Place all your Scharine quote PDFs in the folder:
   `C:\Users\bromayo\OneDrive - Tesla\Quotes\Scharine Quotes`

2. Run the script:

```bash
python quote_scraper.py
```

3. The script will:
   - Process all PDF files in the folder
   - Extract the underlined fields (part number, description, quantity, price)
   - Display a preview in the console
   - Save results to two files:
     - `extracted_quotes.xlsx` (Excel format)
     - `extracted_quotes.csv` (CSV format)

## Output Format

The output files will contain the following columns:
- **source_file**: Name of the PDF file
- **part_number**: Item part number (e.g., F7300K-ZNC)
- **description**: Item description
- **quantity**: Quantity ordered
- **price**: Unit price

## Example Output

```
source_file          part_number      description                           quantity  price
Quote_P034697.pdf    F7300K-ZNC       REAR HITCH, CLR ZINC                  4.0       139.45
Quote_P034697.pdf    F7300H-ZNC       REAR HITCH, CLR ZINC                  16.0      130.95
Quote_P034697.pdf    F518525ASM-C     TOW BAR, EXT. W/BRAKE ASM, CUS        4.0       373.35
```

## Troubleshooting

If items aren't being extracted:
1. Make sure the PDFs are in the correct folder
2. Check that the PDF format matches the Scharine template
3. Try opening the PDF to ensure it's not corrupted
4. Check the console output for any error messages

## Customization

To change the folder path, edit line 104 in `quote_scraper.py`:
```python
folder_path = r"C:\Your\New\Path\Here"
```
