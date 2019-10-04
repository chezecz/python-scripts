import openpyxl

wb = openpyxl.load_workbook('../Stock.xlsx')

ws = wb['ItemStockroomLocationsExport31']

def construct_query(warehouse, location, item, unit_code_2):
	query = f"""update itemloc
				set inv_acct_unit2 = {unit_code_2}
				where item = {item}
				and loc = {location}
				and whse = {warehouse}"""
	print(query)

for i in range(2, ws.max_row):
    construct_query(ws.cell(column = 1, row = i).value, ws.cell(column = 2, row = i).value, ws.cell(column = 3, row = i).value, ws.cell(column = 6, row = i).value)