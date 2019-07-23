def remove_blank_lines(worksheet):
    for i in range(worksheet.max_row, 0, -1):
        #print(worksheet.cell(column = 1, row = i).value == None and worksheet.cell(column = 2, row = i).value == None)
        if worksheet.cell(column = 1, row = i).value == None and worksheet.cell(column = 2, row = i).value == None:
            worksheet.delete_rows(i, 1) 
    return worksheet