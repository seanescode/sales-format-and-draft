import configparser
import os
import ast

import xlwings

from logger import get_logger
from modules.workbook import format_excel
from modules.outlook import create_outlook_email


logger = get_logger()


# find the absolute, full folder path to the exact script file (main.py) that is running right now
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
config_file = os.path.join(SCRIPT_DIR, "config.ini")

config = configparser.ConfigParser()
config.read(config_file)


# Extract data values from the INI file (for email)
recipients = config.get("EMAIL_SETTINGS", "recipients")
cc_recipients = config.get("EMAIL_SETTINGS", "cc_recipients")
email_subject = config.get("EMAIL_SETTINGS", "subject")
email_body = config.get("EMAIL_CONTENT", "body")


# Clean up data structures for your email_tools function requirements
recipient_list = [email.strip() for email in recipients.split(",")]
cc_recipient_list = [email.strip() for email in cc_recipients.split(",")]


# Extract data values from the INI file (for Excel formatting)
bold_title = config.getboolean("EXCEL_REPORT_FORMATTING", "bold_header")
italic_opening_lines = config.getboolean("EXCEL_REPORT_FORMATTING", "italic_first_five_rows")
colour_line_two_to_five = ast.literal_eval(
    config.get("EXCEL_REPORT_FORMATTING", "line_two_to_five_colour")
)
header_colour = ast.literal_eval(
    config.get("EXCEL_REPORT_FORMATTING", "heading_colour")
)
excel_file_location = config.get("EXCEL_REPORT_FORMATTING", "file_location")
excel_sheet_name = config.get("EXCEL_REPORT_FORMATTING", "sheet_name")


app = None
wb = None


try:
    logger.info("Starting sales report automation")

    app = xlwings.App(visible=False)
    wb = app.books.open(excel_file_location)

    logger.info("Excel workbook opened successfully")

    ws = wb.sheets[excel_sheet_name]

    format_excel(
        wb,
        ws,
        header_colour,
        bold_title,
        italic_opening_lines,
        colour_line_two_to_five
    )

    logger.info("Excel formatting completed successfully")

    create_outlook_email(
        recipients=recipient_list,
        cc_recipients=cc_recipient_list,
        subject=email_subject,
        body=email_body,
        attachments=[str(excel_file_location)]
    )

    logger.info("Outlook email draft created successfully")


except Exception as e:
    logger.exception(f"Automation failed: {e}")
    raise


finally:
    if wb is not None:
        wb.close()
        logger.info("Workbook closed")

    if app is not None:
        app.quit()
        logger.info("Excel application closed")

    logger.info("Automation finished")
