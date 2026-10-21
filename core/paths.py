import os
import sys
from pathlib import Path


def user_data_dir():
    if sys.platform == 'win32':
        base = Path(os.environ.get('LOCALAPPDATA', Path.home() / 'AppData' / 'Local'))
    elif sys.platform == 'darwin':
        base = Path.home() / 'Library' / 'Application Support'
    else:
        base = Path(os.environ.get('XDG_CONFIG_HOME', Path.home() / '.config'))
    return base / 'Scraper2GIS'


def default_save_dir():
    return Path.home() / 'Documents' / 'Scraper2GIS'
