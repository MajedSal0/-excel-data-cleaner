# Excel & CSV Data Cleaner

A Streamlit web app that cleans Excel (.xlsx) and CSV files and exports a formatted workbook.

## Run on macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Upload `sample_files/messy_sales.csv` to see a demo. With the three default cleaning options enabled, the sample has 8 input rows, 1 fully empty row removed, 2 duplicate rows removed, and 5 output rows. Missing cells are not fabricated or filled in.

## Features
- Trims leading and trailing whitespace in cells
- Removes completely empty rows
- Removes exact duplicate rows after trimming
- Optional lowercase conversion
- Before/after preview, quality metrics, formatted Excel export with summary sheet

**Limitations:** Duplicate detection uses all columns. No fuzzy matching, validation of business-specific rules, or automatic imputation. Excel input reads the first worksheet. CSV input assumes UTF-8 encoding. Keep client data confidential.
