# Python Excel Automation – Data Cleaning & Reporting

This project demonstrates how Python can automate common Excel tasks that are usually done manually, such as cleaning data and generating summaries.

## What this script does
- Removes duplicate clients (based on email)
- Cleans invalid or missing sales values
- Converts sales data to numeric format
- Generates automatic summaries:
  - Total sales
  - Number of clients
  - Average sales
- Exports a clean Excel report with multiple sheets

## Technologies used
- Python
- pandas
- openpyxl

## Typical use cases
- Sales reports
- Client lists
- Monthly Excel cleanup
- Administrative automation
- Data preparation before analysis

## Project structure
excel_automation_demo/
├── clean_excel.py
├── sample_data.xlsx
├── output/
│ └── summary.xlsx
├── screenshots/
│ └── summary.png
└── README.md


## How to run
1. Install dependencies:
  pip install pandas openpyxl
2. Run the script
   python clean_excel.py
3. Output, The script generates an Excel file at:
   output/summary.xlsx

📌 Author: Boran Gedik
📌 GitHub: https://github.com/borAn77
   
