import configparser
import os
import ast

import xlwings

from logger import get_logger
from modules.workbook import (rename_headings, generate_summary_analytics, find_reporting_start_cell,
                              find_start_cell_subsequent_reports)
from modules.outlook import create_outlook_email

def main():
    logger = get_logger()

    # find the absolute, full folder path to the exact script file (main.py) that is running right now
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(script_dir, "config.ini")

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
    italic_cells = config.getboolean(section="EXCEL_REPORT_FORMATTING", option="italic_font")
    # Read the zoom value as an integer
    zoom_percent = config.getint(section="EXCEL_REPORT_FORMATTING", option="zoom_percentage")

    header_colour = ast.literal_eval(
        config.get(section="EXCEL_REPORT_FORMATTING", option="heading_colour")
    )
    even_row_colour = ast.literal_eval(
        config.get(section="EXCEL_REPORT_FORMATTING", option="even_row_colour")
    )
    odd_row_colour = ast.literal_eval(
        config.get(section="EXCEL_REPORT_FORMATTING", option="odd_row_colour")
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

        rename_headings(ws)

        employee_sales, sales_by_pay_type, sales_by_product = generate_summary_analytics(
            attachment_paths[0],
            excel_sheet_name
        )

        reports_start_cell = find_reporting_start_cell(worksheet=ws)
        second_report_start_cell = find_start_cell_subsequent_reports(worksheet=ws, cell_from=reports_start_cell)
        third_report_start_cell = find_start_cell_subsequent_reports(worksheet=ws, cell_from=second_report_start_cell)

        reports_start_cell.options(index=False).value = employee_sales
        second_report_start_cell.options(index=False).value = sales_by_product
        third_report_start_cell.options(index=False).value = sales_by_pay_type


        format_excel(
            wb,
            ws,
            header_colour,
            even_row_colour,
            odd_row_colour,
            bold_title,
            italic_cells,
            zoom_percent
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
