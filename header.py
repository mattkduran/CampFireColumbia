#!/usr/bin/env python
from typing import Any, Union

from win32com.client.dynamic import CDispatch

__author__ = "Matt Duran"
__copyright__ = "None"
__credits__ = ["Matt Duran"]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Matt Duran"
__email__ = "matduran@campfirecolumbia.org"
__status__ = "Development"

import gc
import os
import time
import progressbar
import win32com.client as win32
from pathlib import Path
import pandas as pd

# Directories and files for data processing
localUser = Path.home()
workingDir = os.path.join(str(localUser),"Desktop", "Forms Report")
destFile = os.path.join(str(localUser), "Desktop", "Forms Report", "dont_open.xlsx")
PPSsource = os.path.join(str(localUser),"Desktop", "Forms Report", "PPSsource.xlsx")
completedDirPPS = os.path.join(str(localUser),"Desktop", "Forms Report", "Completed", "PPS\\")
WLWVsource = os.path.join(str(localUser), "Desktop", "Forms Report", "WLWVsource.xlsx")
completedDirWLWV = os.path.join(str(localUser), "Desktop", "Forms Report", "Completed", "WLWV\\")

# Arrays of school names for filtering
PPSschools = ['Hayhurst', 'Hollyrood', 'Fernwood', 'Woodlawn',
              'Rose City Park', 'James John', 'Sunnyside',
              'Creative Science', 'Peninsula', 'Other']
WLWVschools = ['Bolton', 'Sunset', 'Willamette', 'Trillium Creek',
               'Cedaroak', 'Stafford']


# Variables for working in Excel
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.ScreenUpdating = False
excel.DisplayAlerts = False
excel.EnableEvents = False


def checkExists():
    checkSum = 0
    if os.path.isfile(destFile):
        checkSum += 1
    else:
        print(destFile, " not found!")
    if os.path.isfile(PPSsource):
        checkSum += 1
    else:
        print(PPSsource, " not found!")
    if os.path.isfile(WLWVsource):
        checkSum += 1
    else:
        print(WLWVsource, " not found!")
    if os.path.exists(completedDirPPS):
        checkSum += 1
    else:
        print(completedDirPPS, " not found!")
    if os.path.exists(completedDirWLWV):
        checkSum += 1
    else:
        print(completedDirWLWV, " not found!")
    return checkSum


# Load main dataframe for PPS from source file
def loadMasterPPS(filename):
    frame = pd.read_excel(filename)
    frame.columns = ['First Name', 'Last Name', 'Birthday', 'Drop1', 'Allergies Answer', 'Drop2', 'Medical Answer',
                     'Drop3', 'Other School Answer', 'Drop4', 'Grade Answer', 'Drop5', 'Main School Answer']
    frame.drop(columns=['Drop1', 'Drop2', 'Drop3', 'Drop4', 'Drop5'])
    frame['Student Name'] = frame['Last Name'] + ", " + frame['First Name']

    masterFrame = pd.DataFrame(frame, columns=['Student Name', 'Birthday', 'Allergies Answer',
                                               'Medical Answer', 'Other School Answer',
                                               'Grade Answer', 'Main School Answer'])
    return masterFrame


# Load main dataframe for WLWV from source file
def loadMasterWLWV(filename):
    frame = pd.read_excel(filename)
    frame.columns = ['First Name', 'Last Name', 'Birthday', 'Drop1', 'Allergies Answer', 'Drop2', 'Medical Answer',
                     'Drop3', 'Grade Answer', 'Drop4', 'Main School Answer']
    frame.drop(columns=['Drop1', 'Drop2', 'Drop3', 'Drop4'])
    frame['Student Name'] = frame['Last Name'] + ", " + frame['First Name']

    masterFrame = pd.DataFrame(frame, columns=['Student Name', 'Birthday', 'Allergies Answer',
                                               'Medical Answer', 'Grade Answer', 'Main School Answer'])
    return masterFrame


# Filter main dataframe for one specific school
def filterframe(masterFrame, name):
    subFrame = masterFrame[masterFrame['Main School Answer'] == name]
    return subFrame


# Split out dataframe for PPS where schools are labeled as Other
def splitFrameOther(modifiedFrame):
    with pd.ExcelWriter(destFile) as writer:
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Grade Answer',
                                                                         'Other School Answer'], sheet_name='Grade')
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Birthday',
                                                                         'Other School Answer'], sheet_name='Birthday')
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Allergies Answer',
                                                                         'Other School Answer'], sheet_name='Allergies')
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Medical Answer',
                                                                         'Other School Answer'], sheet_name='Medical')
        writer.save()
    del modifiedFrame
    return


# Split out dataframe for each individual tab
def splitFrame(modifiedFrame):
    with pd.ExcelWriter(destFile) as writer:
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Grade Answer'],
                               sheet_name='Grade')
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Birthday'],
                               sheet_name='Birthday')
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Allergies Answer'],
                               sheet_name='Allergies')
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student Name', 'Medical Answer'],
                               sheet_name='Medical')
        writer.save()
    del modifiedFrame
    return


# Export files out to directory for PPS
def exportPPS(filename, name):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    newName = str(name + timestamp)
    os.chdir(completedDirPPS)
    source = excel.Workbooks.Open(filename)
    source.SaveAs(completedDirPPS+newName)
    os.chdir(workingDir)
    return


# Export files out to directory for WLWV
def exportWLWV(filename, name):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    newName = str(name + timestamp)
    os.chdir(completedDirWLWV)
    source = excel.Workbooks.Open(filename)
    source.SaveAs(completedDirWLWV+newName)
    os.chdir(workingDir)
    return


# Loop for both arrays to process all schools
def looper(PPSmasterFrame, WLWVmasterFrame):
    # Processing PPS
    progressbar.streams.flush()
    progressbar.streams.wrap_stdout()
    print("Processing PPS...")
    with progressbar.ProgressBar(max_value=len(PPSschools)) as bar:
        for i in range(len(PPSschools)):
            schoolName = str(PPSschools[i])
            schoolFrame = filterframe(PPSmasterFrame, schoolName)
            if(schoolName != "Other"):
                splitFrame(schoolFrame)
            else:
                splitFrameOther(schoolFrame)
            exportPPS(destFile, schoolName)
            bar.update(i)

    # Processing WLWV
    print("Processing WLWV...")
    with progressbar.ProgressBar(max_value=len(WLWVschools)) as bar:
        for i in range(len(WLWVschools)):
            schoolName = str(WLWVschools[i])
            schoolFrame = filterframe(WLWVmasterFrame, schoolName)
            splitFrame(schoolFrame)
            exportWLWV(destFile, schoolName)
            bar.update(i)
    return


def destroy(df1, df2):
    global workingDir
    del workingDir
    global destFile
    del destFile
    global PPSsource
    del PPSsource
    global WLWVsource
    del WLWVsource
    global completedDirWLWV
    del completedDirWLWV
    global completedDirPPS
    del completedDirPPS
    del df1
    del df2
    global PPSschools
    del PPSschools[:]
    global WLWVschools
    del WLWVschools[:]
    global excel
    excel.Application.Quit()
    del excel
    gc.collect()
    return