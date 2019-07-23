import os
import shutil
import glob

from config.parse_config import parse_config
from utilities.relative_to_absolute import update_path

config = parse_config()
config = config["FILEPATH"]

def check_file():
    filepath = config["jca_report"]
    filepath_output = update_path(config["excel"])
    for file in glob.glob(filepath):
        shutil.move(file, filepath_output)
    if os.path.exists(filepath_output):
        return True