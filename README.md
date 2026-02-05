# ExcelViewer V.1

A simple desktop application for viewing and filtering Excel files.  
Built with Python, Pandas, and Tkinter.

This tool is made for quick data inspection without opening Microsoft Excel.

## Features
- Open `.xlsx` and `.xls` files
- View data in a table
- Sort data by date/time column
- Filter data by keyword
- Select column to filter
- Live search while typing
- Auto-generate `fullname` from first and last name

## Requirements
- Python 3.9+
- pandas
- openpyxl

## How to Run
1. Clone or download this repository
2. Install dependencies
3. Run:
```bash
python main.py
```
4. Click **Open Excel**
5. Select an Excel file
6. Choose a column and start typing to filter

## Excel File Format
The Excel file must:
- Have a header row
- Use exact column names (case-sensitive)
- Start data from the row below headers

Required columns:
- `firstname`
- `lastname`
- `occurred_at` (used for sorting)

The app automatically creates:
```bash
fullname = firstname + " " + lastname
```
