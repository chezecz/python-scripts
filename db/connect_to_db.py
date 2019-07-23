import pyodbc
import base64

from config.parse_config import parse_config

def get_config(role):
    config = parse_config()
    role = 'sql_' + role
    config = config[role]
    server = config['server']
    database = config['database']
    username = config['username']
    password = base64.b64decode(config['password']).decode('utf-8')
    driver = config['driver']
    
    return {'server': server, 'database': database, 'username': username, 'password': password, 'driver': driver}

def connect_to_db(query, role):
    config = get_config(role)
    connect_string = f"DRIVER={config['driver']};SERVER={config['server']};DATABASE={config['database']};UID={config['username']};PWD={config['password']};"
    with pyodbc.connect(connect_string) as conn:
        with conn.cursor() as cursor:
            cursor.execute(query)
            result = cursor.fetchall()
            conn.commit()
            return result, cursor.description
    