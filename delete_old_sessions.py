from db.connect_to_db import connect_to_db

def get_sessions():
    query = """select * from ConnectionInformation
                where UserName not like '$%'
                and UserName <> 'sa'
                and RecordDate < CONVERT(VARCHAR(10), getdate(), 111)
                """
    
    return connect_to_db(query, 'guest')

def remove_sessions():
    results, _ = get_sessions()
    if results:
        usernames = [(result[1], result[0]) for result in results]
        remove_query(usernames)

def remove_query(usernames):
    for username in usernames:
            query = f"""SET NOCOUNT ON; delete from ConnectionInformation
                where UserName = '{username[0]}'
                and ConnectionID = '{username[1]}'
                select 1
            """
            connect_to_db(query, 'admin')

if __name__ == '__main__':
    remove_sessions()