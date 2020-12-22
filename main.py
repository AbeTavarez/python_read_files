import openpyxl
import os

# os.chdir('/Users/efrentavarez/Downloads')


workbook = openpyxl.load_workbook('example.xlsx')

print(type(workbook))

sheet = workbook.get_sheet_by_name('Sheet1')
# print(sheet)

list_sheet = workbook.sheetnames()
print(list_sheet)
