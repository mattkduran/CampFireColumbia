# Camp Fire Columbia Data Sorter
This program was initially created in order to better sort data exported from Camp Fire's Student Information System (Active) to a CSV format.

Initially, each report was created with multiple columns split out by the following values:
  - First name
  - Last name
  - Birthdate
  - Questions for allergies 
  - Answers for allergies
  - Questions for medical conditions
  - Answers for medical conditions
  - Questions for which school they were attending programs
  - Answers for which school they were attending programs
  - Questions for current grade level
  - Answers for current grade level
  - Questions for the student's main campus
  - Answers for the student's main campus
  
 Initially each report was ran manually -- the SIS would dump the CSV, then each students name was retyped to Last name, First name. Finally, a new workbook in Excel was created for each category (Grade, Allergies, Medical Conditions, Main Campus, Other campus if applicable) for all schools participating in PPS and WLWV.
 
 # Workflow after program
 After the creation of the program, now files are exported from the SIS and saved to a folder as defined in header.py (C:\Users\CURRENT_USER\Desktop\Forms Report). The program was compiled to a single executable for Windows 10, this is ran, and the reports are generated to be sent out in the Completed Directory.
 
```
Folder layout:
    Forms Report:
    |   dont_open.xlsx
    |   PPSsource.xlsx
    |   WLWVsource.xlsx
    |----Completed:
         |-----PPS:
           Completed reports for PPS
         |-----WLWV:
            Completed reports for WLWV
```

# How program works
- Each file is checked to verify that the above structure is intact, if not then the program does not begin
- Once each file is confirmed to exist, PPSsource.xlsx and and WLWVsource.xlsx are both loaded into data frames
- All columns containing questions are dropped and only answers are kept
- Each remaining column is moved to a seperate tab specific to that column (grade level is moved the Grade tab, etc.). This is repated for each school and saved into their own workbooks. 
