import openpyxl

wb = openpyxl.load_workbook('../Stock.xlsx')

ws = wb['ItemStockroomLocationsExport31']

for i in range(2, ws.max_row):
    #construct_sql_query(ws.cell(column = 1, row = i).value, round(ws.cell(column = 7, row = i).value, 4))
    construct_exec_query(ws.cell(column = 1, row = i).value)