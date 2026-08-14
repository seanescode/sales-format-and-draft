import pandas


def rename_headings(worksheet):
    last_column = worksheet.range("A1").end("right").column

    # Start at column 1 and step through to the last column
    for cell in range(1, last_column + 1):
        # Read the value once using clean .cells syntax
        selected_cell = worksheet.cells(1, cell).value

        if selected_cell == "Quantity":
            worksheet.cells(1, cell).value = "Qty"
        elif selected_cell == "Unit Price (€)":
            worksheet.cells(1, cell).value = "Price"
        elif selected_cell == "Discount (%)":
            worksheet.cells(1, cell).value = "Discount"
        elif selected_cell == "Total (€)":
            worksheet.cells(1, cell).value = "Total"
        elif selected_cell == "Payment Method":
            worksheet.cells(1, cell).value = "Pay Type"


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
    return worksheet.range(body_starting_cell).expand('table')


def _format_table_body(worksheet, table_top_left_cell, even_row_color, odd_row_color, italic):
    # format text
    body_range = _get_table_body_range(worksheet, table_top_left_cell)
    body_range.font.italic = italic
    body_range.api.HorizontalAlignment = -4131

    # color rows
    first_cell_row = body_range.row
    first_cell_column = body_range.column

    last_cell_row = worksheet.range((first_cell_row, first_cell_column)).end("down").row
    last_cell_column = worksheet.range((last_cell_row, first_cell_column)).end("right").column


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
    employee_sales = df.groupby('Staff')['Total (€)'].sum().reset_index()
    sales_by_pay_type = df.groupby('Payment Method')['Total (€)'].sum().reset_index()
    sales_by_product = df.groupby('Product')['Total (€)'].sum().reset_index()

    return employee_sales, sales_by_pay_type, sales_by_product


def _find_analytics_start_cell(worksheet, starting_cell):
    start_cell = worksheet.range(starting_cell)
    last_col = start_cell.end("right").column
    return worksheet.cells(start_cell.row, last_col + 2)


def _find_subsequent_analytics_start_cell(worksheet, cell_from):
    bottom_cell = cell_from.end("down")
    bottom_cell_row = bottom_cell.row
    bottom_cell_column = bottom_cell.column
    return worksheet.cells(bottom_cell_row + 2, bottom_cell_column)


def write_analytics(file_path, sheet_name, worksheet, main_table_start_cell):
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
    #Expand columns horizontally so wide text fits
    worksheet.autofit(axis="columns")
    # Expand rows vertically but add 3 points of padding so nothing clips
    worksheet.used_range.rows.autofit()
    for row in worksheet.used_range.rows:
        row.row_height = row.row_height + 3  # Adds tiny safety buffer
