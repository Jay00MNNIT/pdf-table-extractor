# or 📄 PDF Table Extractor

A Python script to extract tables from system-generated PDF files and save them as clean Excel (`.xlsx`) and CSV (`.csv`) files.

---

## 🚀 Features

- Extracts text and organizes it into tabular format based on text position.
- Automatically detects rows using vertical alignment (`top` coordinate).
- Merges nearby words into columns based on gap thresholds.
- Saves output in both Excel and CSV formats.
- Handles numeric conversions for cleaner Excel outputs.
- Works without converting PDF to image, and without using Camelot/Tabula.

---

## 🛠 Requirements

- Python 3.8+
- `pdfplumber`
- `pandas`
- `openpyxl` (for Excel writing)

### Install dependencies:

```bash
pip install -r requirements.txt
```

---

## 📂 How to Use

1. Clone this repository or download the files.
2. Place your input PDF (e.g., `in_scoreme.pdf`) in the project folder.
3. Run the script:

```bash
python convert_pdf.py
```

4. Enter the name of your PDF when prompted.
5. Output files will be saved in the `convert_pdf/` folder.

### 📁 Output Includes:

- Excel file: `convert_pdf/page_X_smart_table.xlsx`
- CSV file: `convert_pdf/page_X_smart_table.csv`

---

## 📌 Notes

- This tool works best on **system-generated PDFs** (not scanned images).
- You may ignore `CropBox missing` warnings, they're safe.
- Output files will only be generated for pages with valid table structures.

---

## 💡 Future Improvements

- Interactive table preview before export.
- PDF page range selection.
- Web-based UI.

---

## 📬 License & Contribution

- Free to use and modify. Contributions are welcome!

---

## 🙌 Acknowledgements

- Built using [pdfplumber](https://github.com/jsvine/pdfplumber)

More upgradation are required for future improvement
Feel free to improve it.

Demo Video ---------> https://drive.google.com/file/d/1c01ZJjtzhBTGeQx_BZb3GGbmUBlOZASS/view?usp=drive_link
