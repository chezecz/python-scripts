def add_new_columns(worksheet, parameter, column_number):
    worksheet.insert_cols(parameter, column_number)
    header_name = 0
    for i in range(parameter, parameter + column_number):
        header_name += 30
        if header_name <= 120:
            worksheet.cell(column = i, row = 1).value = "<=" + str(header_name)
        else:
            worksheet.cell(column = i, row = 1).value = "> 120"
            
    return worksheet