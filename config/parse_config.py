import configparser

from utilities.relative_to_absolute import update_paths

def parse_config():
    config = configparser.ConfigParser()
    config.read(update_paths(['../config.ini', '../user.ini', '../email_settings.ini']))
    return config