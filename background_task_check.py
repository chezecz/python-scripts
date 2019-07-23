from db.connect_to_db import connect_to_db
from utilities.send_yagmail import send_custom_email

def bg_query():
    query = """select DATEDIFF(n,[StartDate],getdate()) as TimeTaken, BGH.TaskName, BGH.RequestingUser, BGH.SubmissionDate, BGH.StartDate, BGH.CompletionDate, BGH.TaskErrorMsg, UN.UserDesc, UN.EmailAddress from BGTaskHistory BGH
LEFT OUTER JOIN Usernames UN
on BGH.RequestingUser = UN.Username
where StartDate >= CONVERT(VARCHAR(10), getdate(), 111)
and (BGH.CompletionDate is NULL and DATEDIFF(n,[StartDate],getdate()) > 50
or BGH.TaskErrorMsg IS NOT NULL)
                """
    results, _ = connect_to_db(query, 'guest')
    send_message(results)

def send_message(results):
    messages = []
    content = {
        'to': "",
        'subject': "",
        'content': ""
    }
    if results:
        messages = [(result[1], result[2], result[3], result[6], result[7], result[8]) for result in results]
    for message in messages:
        #content['to'] = f"{message[1]}@lemacaustralia.com.au"
        content['to'] = message[5]
        name_user = message[4].split(" ")[0]
        content['subject'] = f"The task {message[0]} takes too long to complete or completed with an error. Please have a look!"
        body_message = f"Hi {name_user}, <br><br>The task {message[0]} started at {message[2]} "
        if message[3]:
            content['subject'] = f"The task {message[0]} completed with an error. Please have a look!"
            body_error = f" completed with the error: {message[3]}"
            body_message += body_error
        else:
            body_message += " took too long to complete"
        body_message += "<br><br>Regards,<br>Syteline Admin"
        content['content'] = body_message
        print(content)
        #send_custom_email(content)

if __name__ == '__main__':
    bg_query()