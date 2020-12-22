import openpyxl
import os

# User if file is in a different directory
os.chdir('/Users/<username>/Downloads')

# Load workbook by name
workbook = openpyxl.load_workbook('example.xlsx')

# Shows type
print(type(workbook))

# Grabs sheet by name
sheet = workbook['Sheet1']
print(sheet)

# returns array of sheets
list_sheet = workbook.sheetnames
print(list_sheet)

# Gets cell from a sheet
print(sheet['A1'])
