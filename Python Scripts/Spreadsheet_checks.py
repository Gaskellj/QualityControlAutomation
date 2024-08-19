"""
Authored by James Gaskell

08/12/2024

Edited by:

"""

from pypdf import PdfReader
import numpy as np
import pandas as pd
import tkinter as tk
import time, sys, random, openpyxl, os, Date_formatter, Excel_reader_writer, Location_checker


"""Print slow function to simulate typing
    makes interfce more friendly but not necessarily needed. Typing speed changed by altering typing_speed to reflect words per minute
Args:
    String t: The string typed out
"""
def print_slow(t):
    typing_speed = 60 #wpm
    for l in t:
        sys.stdout.write(l)
        sys.stdout.flush()
        time.sleep(random.random()*10.0/typing_speed)
    print ('')

"""Drops the uneccesary fields from the dataframe to make it easier to use
    Can be expanded - only the fields needed by the program should be in the dataframe since we read the full df again before writing out
Args:
    Dict: dataframes for each sheet
    Keys: keys for the dataframes dictionary - in this case sheet names
Returns:
    Dict: dataframes without the unecessary fields
"""
def clean_dataFrames(dfs, keys):
    for key in keys:
        dfs[key].drop(['documents','type','local_identifier','label (title)','creator_role','contributor_role','genre'], axis=1, inplace=True) #Can be taken out before final version, just helps to view simpler spreadsheet
    return(dfs)


"""Resets the color of rows in the worksheet so they can be reassessed based on the progra's findings
Args:
    worksheet: current sheet opened as part of a Workbook with the openpyxl library
    Int: i refering to the row index after opening the worksheet
    Dict: the colors that are used by the program to highlight errors
        This way we can avoid removing colorson the spreadsheet not caused by previous program runs

"""
def reset_color_rows(ws, i, error_fills):
    fill_reset = openpyxl.styles.PatternFill(fill_type=None)
    if ws.cell(row=i+2, column=1).fill.start_color.index in error_fills.values():
        for y in range(1, ws.max_column+1):
            ws.cell(row=i+2, column=y).fill = fill_reset


"""Identifies the problem rows in the spreadsheet and highlights them depending on the error type
Args:
    String: sheetname - the current sheet being worked on
    Array[String]: the current problem rows, identified by their filename
    String: the type of error currently being added to the spreadsheet - this dictates highlight color
    Boolean: informs the subroutine wether a color reset is required for the spreadsheet - runs when the first error highlight is being added
Returns:
    Boolean: depending on wether the process was successful or not. Fails usually occur due to the excel spreadsheet being open in another application
"""
def identify_problem_rows(sheetname, problem_rows, type, reset_rows):

    #type will be used to determine if the passed problem rows are filename, date format etc.

    error_fills = {"Filename":"FFFFADB0", "Duplicate":"FFADD8E6", "DateFormat":"FFFFFFC5"}

    if type in error_fills.keys():
        fill_hex = error_fills[type]
    else:
        fill_hex = type

    err_fill = openpyxl.styles.PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")

    xl_file = pd.ExcelFile(filepath)       
    dt = pd.read_excel(xl_file, sheetname)

    wb = openpyxl.load_workbook(filepath)
    ws = wb[sheetname]

    for index, row in dt.iterrows():

        if reset_rows:
            reset_color_rows(ws, index, error_fills)

        for p in problem_rows:
            #print(dt['Filename'][index])
            if p == dt['Filename'][index]:
                for y in range(1, ws.max_column+1):
                    ws.cell(row=index+2, column=y).fill = err_fill
    
    try: 
        wb.save(filepath)
    except:
        wb.close()
        xl_file.close()
        return False

    wb.close()
    xl_file.close()
    
    return True

"""Creates the report text files containing errors
Args:
    String: sheetName denoting which sheet is currently being worked on - Box 5, Box 6 etc.
    Array[String]: reportMajor containing the error messages ad filenames to be outputted in the report
    Array[String]: reportNoLocation containing minor location errors for the report
    Array[String]: reportMinor will be used for other errors that aren't in these categories
"""   
def createReportText(sheetName, reportMajor, reportNoLocation, reportMinor):

    newpath = (os.path.dirname(filepath) + "\\Reports\\")
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    outputFile = open(newpath + sheetName + r" Report.txt", "w+")
    outputFile.write("------------Major Issues------------\n\n")
    for error in reportMajor:
        outputFile.write(error + "\n")
    outputFile.write("\n\n\n\n------------Minor Issues------------\n\n")
    for error in reportNoLocation:
        outputFile.write(error + "\n")
    for error in reportMinor:
        outputFile.write(error + "\n")
    outputFile.close()


"""The sheet loop that checks each sheet for errors and creates the error reports
    Contains error checking for the whole program - if any process is not successful and is unhandled an error message is produced
Args:
    Dict: dataframes - all the dataframes in the document split by sheet
    Array[String]: the names of the sheets in the document
"""
def SheetLoop(dfs, sheets):

    for sheet in sheets:

        reportNoLocation = [] # These must be in this subroutine loop to prevent report data carrying over to otherr boxes
        reportMajor = []
        reportMinor = []
        location_problem_rows = []
        duplicate_problem_rows = []
        date_problem_rows = []
        sheet_success = []

        success = Date_formatter.check_date_format(dfs[sheet], reportMajor, date_problem_rows)

        if success:
            sheet_success.append(identify_problem_rows(sheet, date_problem_rows, "DateFormat", True))
            success = Location_checker.check_location_filename(dfs[sheet], reportMajor, reportNoLocation, location_problem_rows) #First check for location/filename discrepancies
        if success:
            sheet_success.append(identify_problem_rows(sheet,location_problem_rows,"Filename", False)) #Colors the problem rows in red for location/filename problems, resets rows from previous runs of the program
            success = Location_checker.check_duplicate_filenames(dfs[sheet], reportMajor, duplicate_problem_rows) #Second check - duplicate filenames
        if success:
            sheet_success.append(identify_problem_rows(sheet, duplicate_problem_rows, "Duplicate", False)) #Colors the problem rows in blue for duplicates, doesn't reset rows as it is the scond step
        if False not in sheet_success:
            createReportText(sheet, reportMajor, reportNoLocation, reportMinor)
            error_rate = round((len(reportMajor)/ (dfs[sheet].shape[0]) * 100) , 1)
            print(sheet + " Completed.     Error Rate: " + str(round(error_rate,1)) + "%")
        else:
            break

    if False in sheet_success:
        tk.messagebox.showerror("File error", "Save Failed. Please ensure the excel file is closed and try again")
    else:
        tk.messagebox.showinfo("Success", "Success! Please check the excel file for issues")


"""Begins the sheetloop that systematically checks each sheet
"""
def run_checks():
    global filepath

    reader_writer = Excel_reader_writer

    try:
        filepath = reader_writer.get_file()

        dfs, sheets = reader_writer.get_dataFrames(filepath)
        print_slow("File loaded...")
        dfs = clean_dataFrames(dfs, sheets)

        #check_location_filename(dfs['Box 6'], reportMajor, reportNoLocation) test
        SheetLoop(dfs, sheets)
    except ValueError:
        tk.messagebox.showinfo("Error", "No File was selected")


    ### For Debugging purposes take out try except

# run_checks()    

# Need to figure out a way to allow leading and trailing 0s for filenames

