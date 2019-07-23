def get_sum(worksheet):
	
    max_row = worksheet.max_row
    st_cell1 = worksheet.cell(column = 6, row = 2)
    st_cell2 = worksheet.cell(column = 7, row = 2)
    f_cell1 = worksheet.cell(column = 6, row = max_row)
    f_cell2 = worksheet.cell(column = 7, row = max_row)
    worksheet.cell(column = 6, row = max_row+1).value = f'= SUM({st_cell1.coordinate} : {f_cell1.coordinate})'
    worksheet.cell(column = 7, row = max_row+1).value = f'= SUM({st_cell2.coordinate} : {f_cell2.coordinate})'
    worksheet.cell(column = 6, row = max_row+1).number_format = '"$"#,##0.00;[Red]\-"$"#,##0.00'
    worksheet.cell(column = 7, row = max_row+1).number_format = '"$"#,##0.00;[Red]\-"$"#,##0.00'
    
def get_column_sum(worksheet, column, st_row, max_row):
    s_cell = worksheet.cell(column = column, row = st_row)
    f_cell = worksheet.cell(column = column, row = max_row)
    cell_value = f'= SUM({s_cell.coordinate} : {f_cell.coordinate})'
    return cell_value
    
def get_total_sum(worksheet, column, max_row):
    cells_sum = []
    for i in range (2, worksheet.max_row - 1):
        worksheet.cell(row = i, column = column+4).number_format = '#,##0.00'
        worksheet.cell(row = i, column = column+6).number_format = '#,##0.00'
        if worksheet.cell(row = i, column = 1).value and not worksheet.cell(row = i, column = 3).value:
            worksheet.cell(row = i, column = column).number_format = '#,##0.00'
            cells_sum.append(worksheet.cell(row = i, column = column).coordinate)
        if worksheet.cell(row = i, column = 1).value is None:
            worksheet.cell(row = i, column = 1).value = worksheet.cell(row = i + 1, column = 1).value
    return ",".join(cells_sum)