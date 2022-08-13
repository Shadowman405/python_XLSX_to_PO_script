import openpyxl
import zipfile
import re

def parseXLSX(workbook_name, product_code_string):
    file_name = ['ru', 'en', 'ar', 'cs', 'de', 'es', 'fi', 'fr', 'it', 'ja', 'pl', 'pt_br', 'th', 'tr', 'vi', 'zh_cn',
                 'zh_tw', 'ko']
    workbook = openpyxl.load_workbook(workbook_name)
    product_code = product_code_string
    names_array = []
    description_array = []
    package_content_array = []
    #sheets = workbook.sheetnames
    sheet = workbook.active
    counter = 0

    for row_1 in sheet['B2':'S2']:
        for cell in row_1:
            names_array.append(str(cell.value))

    for row_2 in sheet['B3':'S3']:
        for cell in row_2:
            description_array.append(str(cell.value))

    for row_3 in sheet['B4':'S4']:
        for cell in row_3:
            string = str(cell.value)
            string = string.replace("\r", "")
            string = string.replace("\n", "")
            new_str = ' '.join(f'<li>{word}</li>' for word in string.split())
            package_content_array.append(new_str)

    for file in file_name:
        with open(r'blueprint.po', "r+", encoding="utf-8") as f:
            read_data = f.read()
        my_file = open(f"{file}.po", "w+", encoding='utf-8')
        read_data = read_data.replace('%product_code%', product_code)
        read_data = read_data.replace('%product_name%', names_array[counter])
        read_data = read_data.replace('%product_description%', description_array[counter])
        read_data = read_data.replace('%package_content%', package_content_array[counter])
        my_file.write(read_data)
        my_file.close()
        counter += 1

def create_zip(product_code):
    with zipfile.ZipFile(f'{product_code}.zip', mode='w') as archive:
        filenames = ['ru', 'en', 'ar', 'cs', 'de', 'es', 'fi', 'fr', 'it', 'ja', 'pl', 'pt_br', 'th', 'tr', 'vi',
                     'zh_cn',
                     'zh_tw', 'ko']
        for filename in filenames:
            archive.write(f'{filename}.po')

#parseXLSX('Excelsior_pckg_content.xlsx', 'test_new_1')