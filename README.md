# Sales Report Automation

## Overview

This project automates the preparation of a weekly sales report.

The automation:
- Opens the Excel sales report
- Applies the required formatting
- Saves the updated workbook
- Creates an Outlook email draft with the report attached

The goal is to reduce manual formatting and email preparation.

---

## Requirements

The application requires:

- Python installed
- Microsoft Excel installed
- Microsoft Outlook desktop installed

Python packages required:

- xlwings
- pywin32

Install required packages using:

pip install -r requirements.txt

---

## Configuration

The automation settings are stored in:

config.ini

The configuration file controls:

- Excel file location
- Worksheet name
- Email recipients
- Email subject
- Email body
- Report formatting options

Update the configuration file if these details change.

---

## Running the Automation

Run the application using:

python main.py

The automation will:
1. Open the Excel report
2. Apply the formatting changes
3. Save the workbook
4. Create an Outlook email draft

The email can then be reviewed and sent.

---

## Troubleshooting

### Outlook error

Make sure Microsoft Outlook desktop is open before running the automation.

### Excel file not found

Check that the Excel file location in `config.ini` is correct.

### Configuration error

Check that all required settings exist in `config.ini`.