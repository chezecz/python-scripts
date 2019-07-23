from excel.copy_to_workbook import insert_data
from excel.new_workbook import create_new_workbook, save_new_workbook
from excel.sum_excel import get_sum

from utilities.send_yagmail import send_email
from utilities.get_dates import set_date_weekly, get_dates_today
from utilities.relative_to_absolute import update_path

from db.connect_to_db import connect_to_db

def get_query():
    select_query = """select i.u_m, il.whse as Warehouse, il.loc as Location, w.name as 'Warehouse Name', i.item, iw.qty_on_hand as 'Warehouse On Hand', il.qty_on_hand as 'Location On Hand', r.qty_sum as 'Item On Hand', i.description as 'Item Description', il.NewRank as Rank, l.description as 'Location Description', l.loc_type as 'Location Type'  
    from itemloc il
    left outer join itemwhse iw
    on il.item = iw.item
    AND il.whse = iw.whse
    left outer join item i
    on il.item = i.item
    left outer join whse w
    on il.whse = w.whse
    left outer join location l
    on il.loc = l.loc
    left outer join (select il.item, il.loc, il.whse, (SELECT SUM (itemwhse.qty_on_hand)FROM itemwhse WHERE itemwhse.item = il.item) AS qty_sum from itemloc il
    left outer join itemwhse iw
    on il.item = iw.item
    AND il.whse = iw.whse
    where (SELECT SUM (itemwhse.qty_on_hand) as sum_qty FROM itemwhse WHERE itemwhse.item = il.item) < 0
    group by il.item, il.loc, il.whse) r
    on r.item = il.item
    and r.loc = il.loc
    and r.whse = il.whse
    where r.qty_sum < 0
    order by il.whse, il.item, il.NewRank
    """
    
    return select_query
    
if __name__ == '__main__':
    result, headlines = connect_to_db(get_query(), 'guest')
    if result:
        filename = f"../../negative_stock_report_{get_dates_today()}.xlsx"
        filename = update_path(filename)
        workbook, filename = create_new_workbook(filename, 'ItemStockRoomLocation')
        worksheet = insert_data(workbook.worksheets[0], result, headlines)
        get_sum(worksheet)
        save_new_workbook(workbook, filename)
    send_email(set_date_weekly(), 'Larry')
