import time

from utilities.get_task import task_query
from db.connect_to_db import connect_to_db
from utilities.check_file import check_file


def query_report():
    connect_to_db(task_query(), 'admin')
    while not check_file():
        time.sleep(10)
    
if __name__ == '__main__':
    query_report()