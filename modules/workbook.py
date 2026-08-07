def color_header(worksheet, header_color):
    # Start at A1 and find the right-most column
    last_column = worksheet.range("A1").end("right").column
    # Select from A1 to the last column and color it all at once
    worksheet.range((1, 1), (1, last_column)).color = header_color


def format_excel(workbook,
                 worksheet,
                 header_colour,
                 bold_title,
                 italic_title,
                 ):

    color_header(worksheet, header_colour)
    worksheet.range("1:1").font.bold = bold_title
    worksheet.range("1:1").font.italic = italic_title

    worksheet.autofit()
    workbook.save()

