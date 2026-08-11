import configparser
import os
import ast

import xlwings

from logger import get_logger
from modules.workbook import format_excel
from modules.outlook import create_outlook_email

def main():
    logger = get_logger()

    # find the absolute, full folder path to the exact script file (main.py) that is running right now
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(SCRIPT_DIR, "config.ini")

    config = configparser.ConfigParser()
    config.read(config_file)


    # Extract data values from the INI file (for email)
    recipients = config.get(section="EMAIL_SETTINGS", option="recipients")
    cc_recipients = config.get(section="EMAIL_SETTINGS", option="cc_recipients")
    email_subject = config.get(section="EMAIL_SETTINGS", option="subject")
    email_body = config.get(section="EMAIL_CONTENT", option="body")


    # Clean up data structures for your email_tools function requirements
    recipient_list = [email.strip() for email in recipients.split(";")]
    cc_recipient_list = [email.strip() for email in cc_recipients.split(";")]


    # Extract data values from the INI file (for Excel formatting)
    bold_title = config.getboolean(section="EXCEL_REPORT_FORMATTING", option="bold_header")
    italic_header = config.getboolean(section="EXCEL_REPORT_FORMATTING", option="italic_font")
    header_colour = ast.literal_eval(
        config.get(section="EXCEL_REPORT_FORMATTING", option="heading_colour")
    )

    excel_sheet_name = config.get(section="EXCEL_REPORT_FORMATTING", option="sheet_name")
    attachment_paths = config.get(section="EXCEL_REPORT_FORMATTING", option="attachment_paths").split(";")
    attachment_paths = [file.strip() for file in attachment_paths]

    app = None
    wb = None


    try:
        logger.info("Starting sales report automation")

        app = xlwings.App(visible=False)
        wb = app.books.open(attachment_paths[0])

        logger.info("Excel workbook opened successfully")

        ws = wb.sheets[excel_sheet_name]

        format_excel(
            wb,
            ws,
            header_colour,
            bold_title,
            italic_header
        )

        logger.info("Excel formatting completed successfully")

        create_outlook_email(
            recipients=recipient_list,
            cc_recipients=cc_recipient_list,
            subject=email_subject,
            body=email_body,
            attachments=attachment_paths
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

if __name__ == "__main__":
    main()
