from openpyxl.styles import Font

def change_font(worksheet):

    ft = Font(name = 'Arial', size = 10)
    
    for rows in worksheet:
        for cell in rows:
            cell.font = ft
            
    return worksheet
    
def change_format(worksheet, start_column, end_column):
    
    for i in range(2, worksheet.max_row):
        for j in range(start_column, end_column):
            worksheet.cell(row = i, column = j).number_format = '#,##0.00'
            
    return worksheet