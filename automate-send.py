import os
import openpyxl
import win32com.client as win32


# Get the current working directory
cwd = os.getcwd()

# Load the Excel workbook
workbook = openpyxl.load_workbook(os.path.join(cwd, "source.xlsx"))