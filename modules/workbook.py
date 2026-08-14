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


def get_table_header_range(worksheet, starting_cell):
    start_cell = worksheet.range(starting_cell)
    row = start_cell.row
    start_col = start_cell.column
    end_col = start_cell.end("right").column
    return worksheet.range((row, start_col), (row, end_col))

def format_table_header(worksheet, start_cell, color_header, bold_header):
    header_range = get_table_header_range(worksheet=worksheet, starting_cell=start_cell)
    header_range.font.bold = bold_header
    header_range.color = color_header
    header_range.api.Borders(9).LineStyle = 1  # bottom border
    header_range.api.HorizontalAlignment = -4131  # left align text


def get_body_range(worksheet, header_starting_cell):
    header_start_row = worksheet.range(header_starting_cell).row
    header_start_column = worksheet.range(header_starting_cell).column
    header_starting_cell = worksheet.cells(header_start_row + 1, header_start_column)
    return worksheet.range(header_starting_cell).expand('table')



def format_body(worksheet, starting_cell, even_row_color, odd_row_color, italic):
    # format text
    body_range = get_body_range(worksheet, starting_cell)
    body_range.font.italic = italic  # italic body rows
    body_range.api.HorizontalAlignment = -4131  # left align text

    # color rows
    top_left_cell = body_range.cell(1,1) # top left cell of a specific table
    first_cell_row = top_left_cell.row
    first_cell_column = top_left_cell.column

    last_cell_row = worksheet.range((first_cell_row, first_cell_column)).end("down").row
    last_cell_column = worksheet.range((last_cell_row, first_cell_column)).end("right").column


    for row in range(first_cell_row, last_cell_row + 1):
        row_color = even_row_color if row % 2 == 0 else odd_row_color
        worksheet.range((row, first_cell_column), (row, last_cell_column)).color = row_color


def format_table(worksheet,
                 starting_cell,
                 color_header,
                 is_bold_header,
                 even_row_color,
                 odd_row_color,
                 is_italic
                 ):
    format_table_header(worksheet, starting_cell, color_header, is_bold_header)
    format_body(worksheet, starting_cell, even_row_color, odd_row_color, is_italic)


def generate_summary_analytics(file_path, sheet_name):
    df = pandas.read_excel(file_path, sheet_name)
    employee_sales = df.groupby('Staff')['Total (€)'].sum().reset_index()
    sales_by_pay_type = df.groupby('Payment Method')['Total (€)'].sum().reset_index()
    sales_by_product = df.groupby('Product')['Total (€)'].sum().reset_index()

    return employee_sales, sales_by_pay_type, sales_by_product


def find_reporting_start_cell(worksheet, starting_cell):
    start_cell = worksheet.range(starting_cell)
    last_col = start_cell.end("right").column
    return worksheet.cells(start_cell.row, last_col + 2)


def find_start_cell_subsequent_reports(worksheet, cell_from):
    bottom_cell = cell_from.end("down")
    bottom_cell_row = bottom_cell.row
    bottom_cell_column = bottom_cell.column
    return worksheet.cells(bottom_cell_row + 2, bottom_cell_column)


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


def save_workbook(workbook, file_path):
    workbook.save(file_path) # "Save As" a new file
    #r"C:\Users\seane\Documents\bakery_sales_overwrite_file.xlsx"

