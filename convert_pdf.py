import pdfplumber
import pandas as pd
from collections import defaultdict
import os 
import re
import warnings
import logging
import subprocess

warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)

def is_num(s):
    try:
        float(s.replace(",", ""))
        return True
    except ValueError:
        return False
    
# get pdf filename
pdf_path = input("Enter the name of the PDF file (with extension): ").strip()

if not os.path.isfile(pdf_path):
    print(f"File not found: {pdf_path}")
    exit()

output_dir = "convert_pdf"
os.makedirs(output_dir, exist_ok=True)

# space between columns
gap_threshold = 10

# process PDF 
with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        words = page.extract_words()
        row_dict = defaultdict(list)

        for word in words:
            y_rounded = round(word['top'], 1)
            row_dict[y_rounded].append(word)

        sorted_rows = sorted(row_dict.items())

        smart_table = []
        max_cols = 0

        for y, words_in_row in sorted_rows:
            sorted_words = sorted(words_in_row, key=lambda w: w["x0"])
            row = []
            if not sorted_words:
                continue

            current_cell = sorted_words[0]['text'].strip()
            last_x1 = sorted_words[0]['x1']

            for word in sorted_words[1:]:
                gap = word['x0'] - last_x1
                if gap > gap_threshold:
                    row.append(current_cell.strip())
                    current_cell = word['text']
                else:
                    current_cell += " " + word['text']
                last_x1 = word['x1']

            row.append(current_cell.strip())
            if any(cell.strip() for cell in row):
                smart_table.append(row)
                max_cols = max(max_cols, len(row))

        # Normalize column counts
        smart_table = [row + [""] * (max_cols - len(row)) for row in smart_table]

        # convert numeric strings to actual numbers
        for r, row in enumerate(smart_table):
            for c, cell in enumerate(row):
                if is_num(cell):
                    smart_table[r][c] = float(cell.replace(",", ""))

        if smart_table:
            header = smart_table[0]
            df = pd.DataFrame(smart_table[1:], columns=header if header else None)

            # save to excel and csv if table found
            output_file = os.path.join(output_dir, f"page_{i+1}_smart_table.xlsx")
            df.to_excel(output_file, index = False, engine='openpyxl')
            print(f" Table from page {i+1} saved to:")
            print(f"   Excel: {output_file}")
            print("Absolute path of saved file:", os.path.abspath(output_file))

            # Also save as csv for text viewing
            csv_file = os.path.join(output_dir, f"page_{i+1}_smart_table.csv")
            df.to_csv(csv_file, index = False)
            print(f"   CSV:   {csv_file}")
        else:
            print(f"No table found on page {i+1}")

    subprocess.Popen(f'explorer "{os.path.abspath(output_dir)}"')
