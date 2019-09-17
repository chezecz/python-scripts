from db.connect_to_db import connect_to_db

def query():
    return 'select username, EmailAddress, UserDesc from usernames where status = 0 and EmailAddress IS NOT NULL'

def update_query(result):
    for i in result:
        if i[1].split('@')[1] == 'lemacaustralia.com.au':
            if i[0][:-1] == i[2].split(" ")[0].lower():
                name = i[1].split(" ")[0][:1].lower() + i[2].split(" ")[1].lower()
            else:
                name = i[1].split('@')[0]
            email_address = name + '@lemacpackaging.com.au'
            query = f"""update usernames
            set EmailAddress = '{email_address}'
            where username = '{i[0]}'"""
            print(query)

if __name__ == '__main__':
    result, _ = connect_to_db(query(), 'admin')
    update_query(result)
    