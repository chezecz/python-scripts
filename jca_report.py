from utilities.query_report import query_report
from utilities.send_yagmail import send_email
from utilities.check_for_new import check_new
from utilities.get_dates import date_no_new_jobs
from utilities.cleanup_files import cleanup_job

from report import get_report

def send_report():
    cleanup_job()
    query_report()
    dc, source = get_report()

    #check for new entries
    if check_new(source):
        send_email(dc, 'report')
    else:
        dc = date_no_new_jobs()
        send_email(dc, 'report')
    
if __name__ == '__main__':
    send_report()