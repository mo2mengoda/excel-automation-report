# Excel Merge & Auto-Report

Turns hours of manual copy-paste into one command. The script merges every Excel/CSV file in a folder (daily sheets, branch reports, exports), cleans the data, and builds a ready-to-send report with a dashboard and charts.

## What it does
1. **Merges** all `.xlsx` / `.xls` / `.csv` files in a folder, and records which file each row came from
2. **Cleans**: normalizes column names, trims spaces, fixes types, drops invalid rows and duplicates
3. **Calculates** totals per row
4. **Builds a report** with these sheets:
   - **Dashboard**: total revenue, orders, units, average order value, top branch and product, plus charts
   - **By Branch**, **By Product**, **Daily**: summary tables
   - **All Data**: the full clean dataset

## Usage
```bash
pip install -r requirements.txt
python make_sample_data.py                                   # demo files
python merge_report.py --input sample_data --out sales_report.xlsx
```

## Tech
Python · pandas · openpyxl (styled sheets + native Excel charts)

> Your files have different columns? Only the `REQUIRED` list and the summary groupings need to change.
