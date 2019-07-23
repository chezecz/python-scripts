import os
import glob

from config.parse_config import parse_config
from utilities.relative_to_absolute import update_path

config = parse_config()
config = config["FILEPATH"]

def clean_output_folder():
    
    filepath = config["jca_report"]

    for file in glob.glob(filepath):
        if os.path.exists(file):
            os.remove(file)

def clean_home_folder():
    
    filepath = update_path("../../*")
    
    filepaths = []
    
    for file in config:
        filepaths.append(update_path(config[file]))
    
    for file in glob.glob(filepath):
        if file in filepaths:
            if os.path.exists(file):
                os.remove(file)
                
def cleanup_job():
    clean_output_folder()
    clean_home_folder()