#!/usr/bin/env python
from header import *
import os


def main():
    checkValue = checkExists()
    if(checkValue == 5):
        PPSmasterFrame = loadMasterPPS(PPSsource)
        WLWVMasterFrame = loadMasterWLWV(WLWVsource)
        looper(PPSmasterFrame, WLWVMasterFrame)
        destroy(PPSmasterFrame, WLWVMasterFrame)
        print("\nJob completed. Quitting...\n")
    else:
        print("Needed files are missing, please confirm that they have been correctly saved.\n")
        print("Quitting...")
    return 0
