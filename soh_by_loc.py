from db.connect_to_db import connect_to_db

from excel.new_workbook import create_new_workbook, save_new_workbook
from excel.copy_to_workbook import insert_data
from excel.excel_fields import insert_section_description

from utilities.send_yagmail import send_email
from utilities.get_dates import set_date_weekly_soh
from utilities.relative_to_absolute import update_path

from excel.add_lines import add_summary_lines, add_grand_total
from excel.excel_styles import change_font, change_format

def get_query():
    select_query = """select i.product_code as 'Product Code', i.item, i.description, il.whse, il.loc, i.unit_cost as 'Unit Cost', il.qty_on_hand as 'Stock On Hand', (i.unit_cost * il.qty_on_hand) as 'Stock Value' from item i
    left outer join itemloc il
    on i.item = il.item
    where il.qty_on_hand <>  0.00
    ORDER BY i.product_code, i.item ASC"""
    
    return select_query

if __name__ == '__main__':
    result, headers = connect_to_db(get_query(), 'guest')
    filename = f"../../SOH_by_Location.xlsx"
    filename = update_path(filename)
    if result:
        workbook, filename = create_new_workbook(filename, "ItemStockroomLocations")
        worksheet = insert_data(workbook.worksheets[0], result, headers)
        worksheet = add_summary_lines(worksheet, 1)
        worksheet = add_grand_total(worksheet)
        worksheet = insert_section_description(worksheet)
        worksheet = change_font(worksheet)
        worksheet = change_format(worksheet, 6, 9)
        save_new_workbook(workbook, filename)
    send_email(set_date_weekly_soh(), 'Lynda')
