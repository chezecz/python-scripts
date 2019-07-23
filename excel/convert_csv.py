import os
import shutil
import glob
import csv
import codecs
from datetime import datetime
from xlsxwriter.workbook import Workbook

cur_date = datetime.today().strftime('%d%m%Y')

for csvfile in glob.glob(os.path.join('..', '*.csv')):
    filename = csvfile[:-4] + '_' + cur_date + '.xlsx'
    workbook = Workbook(filename, {'constant_memory': True, 'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()
    cell_format_number = workbook.add_format()
    cell_format_number.set_num_format('0.00')
    with open(csvfile, 'rt', encoding = 'utf-16') as f:
        reader = csv.reader(f, delimiter = '\t')
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
        workbook.close()
        destination = 'R:\Job Cost Analysis\Reports for Robert'
        shutil.move(filename, destination)