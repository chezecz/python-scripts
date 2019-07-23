import os

dirname = os.path.dirname(__file__)

def update_paths(relative_paths):
    absolute_paths = []
    for filename in relative_paths:
        absolute_paths.append(os.path.abspath(os.path.join(dirname, filename)))
    return absolute_paths
    
def update_path(filename):
    return os.path.abspath(os.path.join(dirname, filename))