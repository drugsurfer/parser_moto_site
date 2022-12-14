"""
This is a setup.py script generated by py2applet

Usage:
    python setup.py py2app

    'packages': ['certifi', 
    'idna', 'pytz', 'requests', 'six',
    'soupsieve', 'urllib3'],
"""

from setuptools import setup

APP = ['window.py']
APP_NAME = 'PitStop Parser'
DATA_FILES = []
OPTIONS = {
    'iconfile': 'icon.icns',
    #'argv_emulation': True,
}

setup(
    app=APP,
    name=APP_NAME, 
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
