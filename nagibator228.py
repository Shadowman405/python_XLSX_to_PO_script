import openpyxl

file_name = ['ru', 'en', 'ar', 'cs', 'de', 'es', 'fi', 'fr', 'it', 'ja', 'pl', 'pt_br', 'th', 'tr', 'vi', 'zh-cn', 'zh_tw', 'ko']
workbook = openpyxl.load_workbook('IS2 (1945) + IS2 SH Bundles.xlsx')

product_name = ''

sheets = workbook.sheetnames
sheet = workbook.active



for row_1 in sheet['B2':'S2']:
    for row_2 in sheet['B3':'S3']:
        string = ''
        for cell in row_1:
            string = string + str(cell.value) + ' \n'
            for file in file_name:
                my_file = open(f"{file}.txt", "w+", encoding='utf-8')
                my_file.write(string)
                my_file.close()
        for cell in row_2:
            string = string + str(cell.value) + ' \n'
            for file in file_name:
                my_file = open(f"{file}.txt", "w+", encoding='utf-8')
                my_file.write(string)
                my_file.close()

for row in sheet.rows:
    for element in range(17):
        string = ''
        string = string + str(row[element].value) + " \n"
        print(string)


tank_name = ''
ver_1 = '4'
ver_2 = '5.5'