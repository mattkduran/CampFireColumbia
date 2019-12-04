from cx_Freeze import setup, Executable
import gc
import os
import time
import win32com.client
import pandas
import progressbar

base = None

executables = [Executable("main.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {
        'packages':packages,
    },
}

setup(
    name = "Datasource",
    options = options,
    version = "1.0.0",
    description = 'Program for sorting data',
    executables = executables
)