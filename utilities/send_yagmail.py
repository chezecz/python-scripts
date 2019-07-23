import yagmail
import base64

from config.parse_config import parse_config

from utilities.relative_to_absolute import update_path

config = parse_config()
outlook_config = config['outlook']
recipients_config = config['recipients_test'] # get test config, would be overwritten later
# email settings

def send_email(content_email, rec):

    recipients = {
        'to': [],
        'cc': []
    }
    recipients_config = config[rec] # comment this line for testing purposes
    subject = content_email['subject']
    content = content_email['content']
    if content_email['filepath_output'] != "":
        attachment = update_path(content_email['filepath_output'])
    else:
        attachment = ""
    yag = yagmail.SMTP(outlook_config['username'],
                        base64.b64decode(outlook_config['password']).decode('utf-8'), 
                        host=outlook_config['host'], 
                        port=outlook_config['port'], 
                        smtp_starttls=True,
                        smtp_ssl=False)
                        
    for item in recipients_config:
        addresses = recipients_config[item].split(',')
        for addresse in addresses:
            recipients[item].append(addresse)

    # sending email (comment yag.send() to perform tests without sending email)
    #print(recipients, subject, content)
    yag.send(to = recipients['to'], subject = subject, cc = recipients['cc'], contents = [content, attachment])

def send_custom_email(email_content):

    yag = yagmail.SMTP(outlook_config['username'],
                        base64.b64decode(outlook_config['password']).decode('utf-8'), 
                        host=outlook_config['host'], 
                        port=outlook_config['port'], 
                        smtp_starttls=True,
                        smtp_ssl=False)

    recipients = {
        'to':[],
        'cc':['syteline_admin@lemacaustralia.com.au']
    }
    recipients['to'] = email_content['to']
    yag.send(to = recipients['to'], subject = email_content['subject'], cc = recipients['cc'], contents = email_content['content'])