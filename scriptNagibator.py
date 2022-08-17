import openpyxl
from parse_def import *

#Enter xlsx filname
xlsx_name = 'test.xlsx'

#Enter product code in list
product_code = ['test']








#DO NOT TOUCH CODE BELOW !!!!
counter = 0
workbook = openpyxl.load_workbook(xlsx_name)
sheets = workbook.sheetnames
workbook_name = xlsx_name

for sheet in sheets:
    parseXLSX(workbook_name,product_code[counter],counter)
    create_zip(product_code[counter])
    counter += 1
