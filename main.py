import openpyxl
import os

# os.chdir('/Users/efrentavarez/Downloads')


workbook = openpyxl.load_workbook('example.xlsx')

print(type(workbook))
