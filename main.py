from tkinter import *
from parse_def import *
import openpyxl
import zipfile

window = Tk()
window.geometry('600x600')
window.title('Welcome fucking slave')

#Adding lbl
label = Label(window, text='Enter excel filename', font=('Arial Bold', 20))
label.grid(column=0, row=0)

label = Label(window, text=' Enter product code(s)', font=('Arial Bold', 20))
label.grid(column=0, row=1)


#Adding txtField
txt = Entry(window, width=15)
txt.grid(column=1,row=0)

txt_2 = Entry(window, width=15)
txt_2.grid(column=1,row=1)

#Adding btn logic
def clicked():
    #Enter xlsx filname
    xlsx_name = '{}.xlsx'.format(txt.get())
    #Enter product code in list
    product_code = ['test']
    print(xlsx_name)
    res_2 = '{}'.format(txt_2.get())
    product_code = res_2.split()
    print(product_code)

    # DO NOT TOUCH CODE BELOW !!!!
    counter = 0
    workbook = openpyxl.load_workbook(str(xlsx_name))
    sheets = workbook.sheetnames
    workbook_name = xlsx_name

    for sheet in sheets:
        parseXLSX(workbook_name, product_code[counter], counter)
        create_zip(product_code[counter])
        counter += 1

    btn.configure(text='Done')


#Adding button
btn = Button(window, text='Generate zip', command=clicked)
btn.grid(column=0, row=8)


window.mainloop()
