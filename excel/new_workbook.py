import openpyxl

def create_new_workbook(filename, title = 'Sheet1'):

    new_workbook = openpyxl.Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = title
    
    new_workbook.save(filename)
    return new_workbook, filename
    
def save_new_workbook(workbook, filename):
    workbook.save(filename)