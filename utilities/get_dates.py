from datetime import datetime, timedelta

from config.parse_config import parse_config

config = parse_config()

content_config = config['email_content']
name = content_config['name']
greeting = content_config['greeting']
ending = content_config['ending']
error_email = content_config['error']

# getting current date
# and the beginning day of the week

current_date = datetime.today()  # today's date
begging_date = timedelta(days = current_date.weekday()) # current weekday (e.g. Monday, Friday)
starting_date = current_date - begging_date # begging of the week

# first and last days of previous week

first_day = current_date - timedelta(weeks = 1) # 1 week before
last_day = current_date - timedelta(days = 1) # yesterday

def set_dates():

    if (current_date.weekday() == 0):
        filepath_output = "../../job_cost_analysis_" + first_day.strftime('%d%m%Y') + " - " + last_day.strftime('%d%m%Y') + ".xlsx"
        content = "{2},<br><br> The JCA Report for period {0} - {1} is attached below.<br><br> {3},<br> {4}".format(first_day.strftime('%d/%m/%Y'), last_day.strftime('%d/%m/%Y'), greeting, ending, name )
        subject = "JCA Report for " + first_day.strftime('%d/%m/%Y') + " - " + last_day.strftime('%d/%m/%Y') 
    elif (current_date.weekday() == 1):
        filepath_output = "../../job_cost_analysis_" + last_day.strftime('%d%m%Y') + ".xlsx"
        content = "{1},<br><br> The JCA Report for {0} is attached below.<br><br> {2},<br> {3}".format(last_day.strftime('%d/%m/%Y'), greeting, ending, name)
        subject = "JCA Report for " + last_day.strftime('%d/%m/%Y') 
    else:
        filepath_output = "../../job_cost_analysis_" + starting_date.strftime('%d%m%Y') + " - " + last_day.strftime('%d%m%Y') + ".xlsx"
        content = "{2},<br><br> The JCA Report for period {0} - {1} is attached below.<br><br> {3},<br> {4}".format(starting_date.strftime('%d/%m/%Y'), last_day.strftime('%d/%m/%Y'), greeting, ending, name )
        subject = "JCA Report for " + starting_date.strftime('%d/%m/%Y') + " - " + last_day.strftime('%d/%m/%Y') 
        
    return {'filepath_output':filepath_output, 'content':content, 'subject':subject}
    
def date_no_new_jobs():
    filepath_output = ""
    content = "{1},<br><br> {0}.<br><br> {2},<br> {3}".format(error_email, greeting, ending, name )
    subject = "JCA Report for " + starting_date.strftime('%d/%m/%Y') + " - " + last_day.strftime('%d/%m/%Y') 
    
    return {'filepath_output':filepath_output, 'content':content, 'subject':subject}
    
def check_week():
    if current_date.weekday() == 0:
        return True
    else: 
        return False
        
def get_dates_today():
    return current_date.strftime('%d%m%Y')

def get_period_dates():
    if (current_date.weekday() == 0):
        return first_day.strftime('%Y-%m-%d'), last_day.strftime('%Y-%m-%d')
    elif (current_date.weekday() == 1):
        return last_day.strftime('%Y-%m-%d'), last_day.strftime('%Y-%m-%d')
    else:
        return starting_date.strftime('%Y-%m-%d'), last_day.strftime('%Y-%m-%d')
    
def set_date_weekly():
    greeting = content_config['recipient']
    filepath_output = "../../negative_stock_report_" + current_date.strftime('%d%m%Y') +".xlsx"
    content = "{1},<br><br> The negative stock report for {0} is attached below.<br><br> {2},<br> {3}".format(current_date.strftime('%d/%m/%Y'), greeting, ending, name )
    subject = "Negative stock report for " + current_date.strftime('%d/%m/%Y')
    return {'filepath_output':filepath_output, 'content':content, 'subject':subject}

def set_date_weekly_soh():
    greeting = content_config['recipient']
    filepath_output = "../../SOH_by_Location.xlsx"
    content = "{1},<br><br> The SOH by location report for {0} is attached below.<br><br> {2},<br> {3}".format(current_date.strftime('%d/%m/%Y'), greeting, ending, name )
    subject = "SOH by location report for " + current_date.strftime('%d/%m/%Y')
    return {'filepath_output':filepath_output, 'content':content, 'subject':subject}
    
def set_date_weekly_aging():
    greeting = content_config['recipient']
    filepath_output = "../../Aging_report.xlsx"
    content = "{1},<br><br> The Aging by Lot report for {0} is attached below.<br><br> {2},<br> {3}".format(current_date.strftime('%d/%m/%Y'), greeting, ending, name )
    subject = "Aging by Lot report for " + current_date.strftime('%d/%m/%Y')
    return {'filepath_output':filepath_output, 'content':content, 'subject':subject}