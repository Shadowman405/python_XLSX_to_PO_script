import openpyxl

file_name = ['ru', 'en', 'ar', 'cs', 'de', 'es', 'fi', 'fr', 'it', 'ja', 'pl', 'pt_br', 'th', 'tr', 'vi', 'zh_cn', 'zh_tw', 'ko']
workbook = openpyxl.load_workbook('')

product_code = 'test_locs'

names_array = []
description_array = []
counter = 0

sheets = workbook.sheetnames
sheet = workbook.active

for row_1 in sheet['B2':'S2']:
    for cell in row_1:
        names_array.append(str(cell.value))

for row_2 in sheet['B3':'S3']:
    for cell in row_2:
        description_array.append(str(cell.value))


for file in file_name:
    with open(r'blueprint.po', "r+", encoding="utf-8") as f:
        read_data = f.read()
    my_file = open(f"{file}.po", "w+", encoding='utf-8')
    read_data = read_data.replace('%product_code%', product_code)
    read_data = read_data.replace('%product_name%', names_array[counter])
    read_data = read_data.replace('%product_description%', description_array[counter])
    my_file.write(read_data)
    my_file.close()
    counter += 1
