from db.connect_to_db import connect_to_db

from excel.new_workbook import create_new_workbook, save_new_workbook
from excel.copy_to_workbook import insert_data
from excel.excel_fields import insert_section_description

from utilities.send_yagmail import send_email
from utilities.get_dates import set_date_weekly_aging
from utilities.relative_to_absolute import update_path
from utilities.time_diff import time_diff_worksheet

from excel.add_lines import add_summary_lines, add_items, calculate_sums_buckets
from excel.remove_lines import remove_blank_lines
from excel.excel_styles import change_font, change_format
from excel.add_columns import add_new_columns

def get_query():
    select_query = """select res.product_code as 'Product Code', res.item, it.description, res.lot, lt.create_date, res.loc, res.total as Qty, it.unit_cost as 'Unit Cost', (res.total * it.unit_cost) as 'Total Value' from (

select i.product_code, i.item, ll.lot, ll.loc, ll.qty_on_hand as total from item i
    inner join lot_loc ll
    on ll.item = i.item
    where ll.lot IN (select lot from lot_loc
						where qty_on_hand <> 0
							)
) as res
inner join item it 
on it.item = res.item
inner join lot lt
on lt.item = res.item AND lt.lot = res.lot
order by res.product_code, res.item, lt.lot DESC"""
    
    return select_query

if __name__ == '__main__':
    result, headers = connect_to_db(get_query(), 'guest')
    filename = f"../../Aging_report.xlsx"
    filename = update_path(filename)
    if result:
        workbook, filename = create_new_workbook(filename, "AgingReport")
        worksheet = insert_data(workbook.worksheets[0], result, headers)
        worksheet = add_new_columns(worksheet, parameter = 7, column_number = 5)
        worksheet = time_diff_worksheet(worksheet, date_column = 5, value_cell = 12)
        worksheet = add_summary_lines(worksheet, 2, title = 'Item Total')
        worksheet = add_items(worksheet)
        worksheet = calculate_sums_buckets(worksheet, 7, 13)
        worksheet = change_font(worksheet)
        worksheet = change_format(worksheet, 7, 14)
        save_new_workbook(workbook, filename)
    send_email(set_date_weekly_aging(), 'Lynda')