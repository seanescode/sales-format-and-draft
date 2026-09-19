import os

import sales_format_and_draft.dialogs as dialogs
import pandas
import xlwings


def check_spreadsheet_ready(spreadsheet_path):
    if not os.path.exists(spreadsheet_path):
        dialogs.show_error_dialog(
            title="Spreadsheet Not Found",
            message=f"The sales spreadsheet could not be found:\n\n{spreadsheet_path}"
        )
        return False
    if _is_spreadsheet_open(spreadsheet_path):
        dialogs.show_error_dialog(
            title="Spreadsheet Already Open",
            message="The sales spreadsheet is already open.\n\n"
                    "Please close it and try again."
        )
        return False
    return True


def _is_spreadsheet_open(spreadsheet_path):
    target_path = os.path.abspath(spreadsheet_path).lower()

    for app in xlwings.apps:
        for book in app.books:
            if os.path.abspath(book.fullname).lower() == target_path:
                return True

    return False


def _get_table_header_range(worksheet, starting_cell):
    start_cell = worksheet.range(starting_cell)
    row = start_cell.row
    start_col = start_cell.column
    end_col = start_cell.end("right").column
    return worksheet.range((row, start_col), (row, end_col))


def _format_table_header(worksheet, start_cell, color_header, bold_header):
    header_range = _get_table_header_range(worksheet=worksheet, starting_cell=start_cell)
    header_range.font.bold = bold_header
    header_range.color = color_header
    header_range.api.Borders(9).LineStyle = 1
    header_range.api.HorizontalAlignment = -4131


def _get_table_body_range(worksheet, table_top_left_cell):
    table_start_row = worksheet.range(table_top_left_cell).row
    table_start_column = worksheet.range(table_top_left_cell).column
    body_starting_cell = worksheet.cells(table_start_row + 1, table_start_column)
    body_range = worksheet.range(body_starting_cell).expand('table')
    
    # Exclude the last row if it's empty (expand('table') can include an extra row)
    last_row = body_range.row + body_range.rows.count - 1
    last_col = body_range.column + body_range.columns.count - 1
    
    # Check if the last row is empty
    last_row_values = worksheet.range((last_row, body_range.column), (last_row, last_col)).value
    if all(v is None or v == '' for v in last_row_values if v is not None):
        # Remove the last empty row from the range
        if body_range.rows.count > 1:
            body_range = worksheet.range(body_range.row, body_range.column, last_row - 1, last_col)
    
    return body_range


def _format_table_body(worksheet, table_top_left_cell, even_row_color, odd_row_color, italic):
    # format text
    body_range = _get_table_body_range(worksheet, table_top_left_cell)
    body_range.font.italic = italic
    body_range.api.HorizontalAlignment = -4131

    # color rows using the actual body range dimensions
    first_cell_row = body_range.row
    first_cell_column = body_range.column
    last_cell_row = body_range.row + body_range.rows.count - 1
    last_cell_column = body_range.column + body_range.columns.count - 1

    for row in range(first_cell_row, last_cell_row + 1):
        row_color = even_row_color if row % 2 == 0 else odd_row_color
        worksheet.range((row, first_cell_column), (row, last_cell_column)).color = row_color


def _format_table(worksheet,
                  starting_cell,
                  color_header,
                  is_bold_header,
                  even_row_color,
                  odd_row_color,
                  is_italic
                  ):
    _format_table_header(worksheet, starting_cell, color_header, is_bold_header)
    _format_table_body(worksheet, starting_cell, even_row_color, odd_row_color, is_italic)


def format_all_tables(worksheet, main_table_start_cell, color_header,
                      is_bold_header, even_row_color, odd_row_color, is_italic):
    analytics_starting_cell = _find_analytics_start_cell(
        worksheet, main_table_start_cell
    )
    second_starting_cell = _find_subsequent_analytics_start_cell(
        worksheet, analytics_starting_cell
    )
    third_starting_cell = _find_subsequent_analytics_start_cell(
        worksheet, second_starting_cell
    )

    for starting_cell in (
            main_table_start_cell,
            analytics_starting_cell,
            second_starting_cell,
            third_starting_cell,
    ):
        _format_table(
            worksheet,
            starting_cell,
            color_header,
            is_bold_header,
            even_row_color,
            odd_row_color,
            is_italic,
        )


def _generate_summary_analytics(file_path, sheet_name):
    df = pandas.read_excel(file_path, sheet_name)

    employee_sales = (df.groupby('Staff')['Total (€)']
                      .sum()
                      .reset_index()
                      .sort_values('Total (€)', ascending=False)
                      )

    sales_by_pay_type = (df.groupby('Pay Type')['Total (€)']
                         .sum()
                         .reset_index()
                         .sort_values('Total (€)', ascending=False)

                         )

    sales_by_product = (df.groupby('Product')['Total (€)']
                        .sum()
                        .reset_index()
                        .sort_values('Total (€)', ascending=False)
                        )

    return employee_sales, sales_by_pay_type, sales_by_product


def _find_analytics_start_cell(worksheet, starting_cell):
    start_cell = worksheet.range(starting_cell)
    last_col = start_cell.end("right").column
    return worksheet.cells(start_cell.row, last_col + 2)


def _find_subsequent_analytics_start_cell(worksheet, cell_from):
    # Find the actual last row with data by checking from the starting cell
    start_cell = worksheet.range(cell_from)
    start_row = start_cell.row
    start_col = start_cell.column
    
    # Go down to find the last row with data
    current_row = start_row
    while True:
        next_row = current_row + 1
        check_cell = worksheet.cells(next_row, start_col)
        if check_cell.value is None or check_cell.value == '':
            break
        current_row = next_row
    
    return worksheet.cells(current_row + 2, start_col)


def _clear_analytics_area(worksheet, main_table_start_cell):
    analytics_starting_cell = _find_analytics_start_cell(worksheet, main_table_start_cell)
    start_row = analytics_starting_cell.row
    start_col = analytics_starting_cell.column
    
    last_row = worksheet.used_range.last_cell.row
    last_col = worksheet.used_range.last_cell.column
    
    if start_row <= last_row and start_col <= last_col:
        clear_range = worksheet.range((start_row, start_col), (last_row, last_col))
        clear_range.clear_contents()


def write_analytics(file_path, sheet_name, worksheet, main_table_start_cell):
    _clear_analytics_area(worksheet, main_table_start_cell)
    
    analytics = _generate_summary_analytics(file_path, sheet_name)
    employee_sales, sales_by_pay_type, sales_by_product = analytics

    analytics_starting_cell = _find_analytics_start_cell(worksheet, main_table_start_cell)
    analytics_starting_cell.options(index=False).value = employee_sales

    second_starting_cell = _find_subsequent_analytics_start_cell(worksheet, analytics_starting_cell)
    second_starting_cell.options(index=False).value = sales_by_pay_type

    third_starting_cell = _find_subsequent_analytics_start_cell(worksheet, second_starting_cell)
    third_starting_cell.options(index=False).value = sales_by_product


def format_worksheet(worksheet, zoom_percentage):
    worksheet.activate()
    worksheet.book.app.api.ActiveWindow.Zoom = zoom_percentage
    worksheet.book.app.api.ActiveWindow.ScrollColumn = 1
    worksheet.range("A1").select()

    # autosize rows and columns
    # Expand columns horizontally so wide text fits
    worksheet.autofit(axis="columns")
    # Expand rows vertically but add 3 points of padding so nothing clips
    worksheet.used_range.rows.autofit()
    for row in worksheet.used_range.rows:
        row.row_height = row.row_height + 5  # Adds tiny safety buffer
    worksheet.range("B:B").api.EntireColumn.Hidden = True
    worksheet.range("D:D").api.EntireColumn.Hidden = True
    worksheet.range("H:H, F:F, M:M").number_format = "#,##0.00"


def rename_headings(worksheet):
    last_column = worksheet.range("A1").end("right").column

    # Start at column 1 and step through to the last column
    for cell in range(1, last_column + 1):
        # Read the value once using clean .cells syntax
        selected_cell = worksheet.cells(1, cell).value

        if selected_cell == "Discount (%)":
            worksheet.cells(1, cell).value = "Disc (%)"
