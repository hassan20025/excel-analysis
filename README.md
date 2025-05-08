# Sales Report Analyzer (Arabic GUI)

This tool helps analyze sales data from Excel files using a GUI built with `tkinter`. It reads invoice data, calculates total commissions for salespeople, and displays the results in a formatted table.

## Features

- Upload one or more Excel files.
- Identify salespeople by code.
- Calculate commission:
  - `bahaa`: 40%
  - Others: 14%
- Display each seller's total amount, client names, phone numbers, and individual net values.
- Supports Arabic text with reshaping and bidi display.

## Technologies Used

- Python
- `tkinter` for GUI
- `pandas` for Excel data handling
- `openpyxl` for reading Excel
- `arabic_reshaper` and `python-bidi` for proper Arabic display

## How to Run

1. Make sure you have the required libraries:

```bash
pip install pandas openpyxl arabic_reshaper python-bidi tabula-py
