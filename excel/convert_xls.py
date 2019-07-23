import xlrd
import openpyxl
import os

from utilities.relative_to_absolute import update_path

# work with Excel 2003 format

def convert_file(filepath_excel, filepath_data):
    filepath_data = update_path(filepath_data)
    filepath_excel = update_path(filepath_excel)
    book = xlrd.open_workbook(filepath_excel)
    index = 0
    nrows, ncols = 0, 0
    while nrows * ncols == 0:
        sheet = book.sheet_by_index(index)
        nrows = sheet.nrows
        ncols = sheet.ncols
        index += 1
        
    book1 = openpyxl.Workbook()
    sheet1 = book1.active
    sheet1.title = "Sheet1"

    for row in range(0, nrows):
        for col in range (0, ncols):
            sheet1.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)
            
    book1.save(filepath_data)
    if os.path.exists(filepath_excel):
        os.remove(filepath_excel)
    return filepath_data