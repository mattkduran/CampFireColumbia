#!/usr/bin/env python
from typing import Any, Union
from win32com.client.dynamic import CDispatch

__author__ = "Matt Duran"
__copyright__ = "None"
__credits__ = ["Matt Duran"]
__license__ = "GPL"
__version__ = "2.0.0"
__maintainer__ = "Matt Duran"
__email__ = "mduran@campfirecolumbia.org"
__status__ = "Development"

import gc
import os
import time
import progressbar
from win32com.client import Dispatch
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
excel = Dispatch('Excel.Application')
excel.Visible = False
excel.ScreenUpdating = False
excel.DisplayAlerts = False
excel.EnableEvents = False

def menu():
    pps = " Run PPS report (1)\n"
    wlwv = "Run WLWV report (2)\n"
    both = "Run both reports (3)\n"
    border = "####################\n"
    checkValue = checkExists()
    if (checkValue == 5):
        print(border, pps, wlwv, both)
    elif (3 < checkValue < 5):
        print(border, pps, wlwv)
    else:
        print("Needed files are missing, please confirm that they have been correctly saved.\n")
        input("Press enter to quit.")
        print("Quitting...")
        return

    choice = input("Enter value: ")
    print(border)
    if (int(choice) == 3):
        runBoth()
        print("\nJob completed. Quitting...\n")
    elif (int(choice) < 3):
        runOne(choice)
        print("\nJob completed. Quitting...\n")
    return

def runOne(choice):
    if (int(choice) == 1):
        PPSmasterFrame = loadMasterPPS(PPSsource)
        looperOne(PPSmasterFrame, choice)
        destroyOne(PPSmasterFrame)
    if (int(choice) == 2):
        WLWVmasterFrame = loadMasterWLWV(WLWVsource)
        looperOne(WLWVmasterFrame, choice)
        destroyOne(WLWVmasterFrame)
    return

def runBoth():
    PPSmasterFrame = loadMasterPPS(PPSsource)
    WLWVMasterFrame = loadMasterWLWV(WLWVsource)
    looperBoth(PPSmasterFrame, WLWVMasterFrame)
    destroyBoth(PPSmasterFrame, WLWVMasterFrame)
    return

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
    frame.columns = ['First Name', 'Last Name', 'Birth Date', 'Drop1', 'Allergies', 'Drop2', 'Medical',
                     'Drop3', 'Other School Answer', 'Drop4', 'Grade', 'Drop5', 'Main School Answer']
    frame.drop(columns=['Drop1', 'Drop2', 'Drop3', 'Drop4', 'Drop5'])
    frame['Student'] = frame['Last Name'] + ", " + frame['First Name']

    frame['Birth Date'] = frame['Birth Date'].dt.strftime('%m/%d/%y')
    frame.fillna('None', inplace=True)

    masterFrame = pd.DataFrame(frame, columns=['Student', 'Birth Date', 'Allergies',
                                               'Medical', 'Other School Answer',
                                               'Grade', 'Main School Answer'])
    return masterFrame


# Load main dataframe for WLWV from source file
def loadMasterWLWV(filename):
    frame = pd.read_excel(filename)
    frame.columns = ['First Name', 'Last Name', 'Birth Date', 'Drop1', 'Allergies', 'Drop2', 'Medical',
                     'Drop3', 'Grade', 'Drop4', 'Main School Answer']
    frame.drop(columns=['Drop1', 'Drop2', 'Drop3', 'Drop4'])
    frame['Student'] = frame['Last Name'] + ", " + frame['First Name']

    frame['Birth Date'] = frame['Birth Date'].dt.strftime('%m/%d/%y')
    frame.fillna('None', inplace=True)

    masterFrame = pd.DataFrame(frame, columns=['Student', 'Birth Date', 'Allergies',
                                               'Medical', 'Grade', 'Main School Answer'])
    return masterFrame


# Filter main dataframe for one specific school
def filterframe(masterFrame, name):
    subFrame = masterFrame[masterFrame['Main School Answer'] == name]
    return subFrame


# Split out dataframe for PPS where schools are labeled as Other
def splitFrameOther(modifiedFrame):
    with pd.ExcelWriter(destFile, engine='xlsxwriter') as writer:
        workbook = writer.book
        borders = workbook.add_format({'border': 1})
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Grade', 'Other School Answer'],
                               sheet_name='Grade')
        worksheet = writer.sheets['Grade']
        worksheet.set_column('A:C', 18, borders)
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Birth Date', 'Other School Answer'],
                               sheet_name='Birthday')
        worksheet = writer.sheets['Birthday']
        worksheet.set_column('A:C', 18, borders)
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Allergies', 'Other School Answer'],
                               sheet_name='Allergies')
        worksheet = writer.sheets['Allergies']
        worksheet.set_column('A:C', 18, borders)
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Medical', 'Other School Answer'],
                               sheet_name='Medical')
        worksheet = writer.sheets['Medical']
        worksheet.set_column('A:C', 18, borders)
        writer.save()
    del modifiedFrame
    return


# Split out dataframe for each individual tab
def splitFrame(modifiedFrame):
    with pd.ExcelWriter(destFile, engine='xlsxwriter') as writer:
        workbook = writer.book
        borders = workbook.add_format({'border': 1})
        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Grade'],
                               sheet_name='Grade')
        worksheet = writer.sheets['Grade']
        worksheet.set_column('A:B', 18, borders)

        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Birth Date'],
                               sheet_name='Birthday')
        worksheet = writer.sheets['Birthday']
        worksheet.set_column('A:B', 18, borders)

        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Allergies'],
                               sheet_name='Allergies')
        worksheet = writer.sheets['Allergies']
        worksheet.set_column('A:B', 18, borders)

        modifiedFrame.to_excel(writer, index=None, header=True, columns=['Student', 'Medical'],
                               sheet_name='Medical')
        worksheet = writer.sheets['Medical']
        worksheet.set_column('A:B', 18, borders)

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
def looperOne(masterFrame, choice):
    # Processing PPS
    progressbar.streams.flush()
    progressbar.streams.wrap_stdout()
    if (int(choice) == 1):
        print("Processing PPS...")
        with progressbar.ProgressBar(max_value=len(PPSschools)) as bar:
            for i in range(len(PPSschools)):
                schoolName = str(PPSschools[i])
                schoolFrame = filterframe(masterFrame, schoolName)
                if(schoolName != "Other"):
                    splitFrame(schoolFrame)
                else:
                    splitFrameOther(schoolFrame)
                exportPPS(destFile, schoolName)
                bar.update(i)

    # Processing WLWV
    if (int(choice) == 2):
        print("Processing WLWV...")
        with progressbar.ProgressBar(max_value=len(WLWVschools)) as bar:
            for i in range(len(WLWVschools)):
                schoolName = str(WLWVschools[i])
                schoolFrame = filterframe(masterFrame, schoolName)
                splitFrame(schoolFrame)
                exportWLWV(destFile, schoolName)
                bar.update(i)
    return

def looperBoth(PPSmasterFrame, WLWVmasterFrame):
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

def destroyOne(df1):
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
    global PPSschools
    del PPSschools[:]
    global WLWVschools
    del WLWVschools[:]
    global excel
    excel.Application.Quit()
    del excel
    gc.collect()
    return

def destroyBoth(df1, df2):
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
