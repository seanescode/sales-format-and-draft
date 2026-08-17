import os

import xlwings
from modules import config, outlook, spreadsheet

def main():

    app = None
    wb = None

    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(script_dir, "config.ini")

    settings = config.load_config(config_file)

    email_settings = config.get_email_settings(settings)
    excel_settings = config.get_excel_settings(settings)

    spreadsheet_path = excel_settings["attachment_paths"][0]

    if not spreadsheet.check_spreadsheet_ready(spreadsheet_path):
        return
    if not outlook.check_outlook_ready(email_settings["recipients"]):
        return

    try:
        app = xlwings.App(visible=False)
        wb = app.books.open(
            spreadsheet_path
        )
        ws = wb.sheets[
            excel_settings["sheet_name"]
        ]

        spreadsheet.rename_headings(ws)
        spreadsheet.write_analytics(
            file_path=excel_settings["attachment_paths"][0],
            sheet_name=excel_settings["sheet_name"],
            worksheet=ws,
            main_table_start_cell="A1"
        )
        spreadsheet.format_all_tables(
            worksheet=ws,
            main_table_start_cell="A1",
            color_header=excel_settings["header_colour"],
            is_bold_header=excel_settings["bold_header"],
            even_row_color=excel_settings["even_row_colour"],
            odd_row_color=excel_settings["odd_row_colour"],
            is_italic=excel_settings["italic_cells"],
        )
        spreadsheet.format_worksheet(
            ws,
            excel_settings["zoom_percent"]
        )

        wb.save(
            excel_settings["attachment_paths"][0]
        )

        outlook.create_outlook_email(
            recipients=email_settings["recipients"],
            cc_recipients=email_settings["cc_recipients"],
            subject=email_settings["subject"],
            body=email_settings["body"],
            attachments=excel_settings["attachment_paths"]
        )

    finally:
        if wb is not None:
            wb.close()
        if app is not None:
            app.quit()

if __name__ == "__main__":
    main()
