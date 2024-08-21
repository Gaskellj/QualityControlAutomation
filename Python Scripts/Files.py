"""
Authored by James Gaskell

08/21/2024

Edited by:

"""

import pandas as pd
import tkinter as tk

class File():

    filepath = None
    extent = None
    max_file_size = None

    def __init__(self, filepath, extent, max_file_size):
        self.filepath = filepath
        self. extent = extent
        self.max_file_size = max_file_size

class ExcelFile(File):

    sheetnames = []
    dataframes = None

    def __init__(self, filepath):
        self.filepath = filepath
    
    def read_data_frames(self):
        try:
            xl_file = pd.ExcelFile(self.filepath)
            self.dataframes = pd.read_excel(xl_file, sheet_name=None)
            self.sheetnames = xl_file.sheet_names
            xl_file.close()
        except:
            tk.messagebox.showinfo("Error", "File error. Please try again")

    