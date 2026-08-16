# Sales Report Automation

## Overview

<img width="1020" height="651" alt="screen_pic_email_automated" src="https://github.com/user-attachments/assets/7bda58fb-9f41-4696-8107-21d8dc28d6e5" />
<img width="689" height="637" alt="screen_pic_updated_data" src="https://github.com/user-attachments/assets/0cfa4ef3-1736-4ec3-8183-64b7bde2bffa" />
<img width="935" height="640" alt="screen_pic_raw_data" src="https://github.com/user-attachments/assets/ee0beaa7-215c-49a6-af35-3592617c64b0" />


This project automates the preparation of a weekly sales report.

The automation:

* Opens an Excel sales report
* Applies the required formatting
* Saves the updated workbook
* Creates a Microsoft Outlook email draft with the report attached

The goal is to reduce manual formatting, improve consistency, and speed up the report preparation process.

---

## Features

* Reads settings from a `config.ini` file
* Automatically formats an Excel worksheet
* Applies company formatting to headers and report sections
* Auto-sizes rows and columns
* Saves the formatted workbook
* Creates an Outlook email draft with the report attached
* Records application activity using log files

---

## Requirements

The application requires:

* Python 3.10 or later
* Microsoft Excel
* Microsoft Outlook (desktop version)

Install the required Python packages:

```bash
pip install -r requirements.txt
```

---

## Configuration

Application settings are stored in:

```text
config.ini
```

The configuration file controls:

* Excel file location
* Worksheet name
* Email recipients
* CC recipients
* Email subject
* Email body
* Excel formatting options

Update `config.ini` whenever these settings need to change.

---

## Project Structure

```text
automation_library/
│
├── main.py
├── config.ini
├── requirements.txt
├── logger.py
├── modules/
│   ├── workbook.py
│   ├── outlook.py
│   └── dialogs.py
└── tests/
```

---

## Technologies

* Python
* xlwings
* pywin32
* pytest
* logging
* configparser

---

## Running the Automation

Run the application from the project directory:

```bash
python main.py
```

The automation will:

1. Open the Excel report.
2. Apply the configured formatting.
3. Save the workbook.
4. Create an Outlook email draft with the formatted report attached.

The email draft can then be reviewed before sending.

---

## Troubleshooting

### Outlook is not open

Ensure the Microsoft Outlook desktop application is running before starting the automation.

### Excel file cannot be found

Verify that the file path in `config.ini` is correct and the workbook exists.

### Configuration error

Ensure all required settings are present in `config.ini`.

### Missing Python packages

If required packages are missing, install them using:

```bash
pip install -r requirements.txt
```
