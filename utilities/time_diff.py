from datetime import datetime

from utilities.get_dates import get_dates_today
from excel.add_lines import from_excel_ordinal

def time_difference(date):
    cur_date = datetime.today()
    return (cur_date - date).days
    
def calculate_time_diff(date):
    result_diff = time_difference(date)
    if result_diff <= 30:
        return 1
    elif result_diff <= 60:
        return 2
    elif result_diff <= 90:
        return 3
    elif result_diff <= 120:
        return 4
    else:
        return 5
        
def time_diff_worksheet(worksheet, date_column, value_cell):
    max_row = worksheet.max_row
    for i in range(2, max_row+1):
        if worksheet.cell(column = date_column, row = i).value:
            offset_column = calculate_time_diff(worksheet.cell(column = date_column, row = i).value)
            worksheet.cell(column = date_column + offset_column + 1, row = i).value = worksheet.cell(column = value_cell, row = i).value
    
    return worksheet