from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
from parse_def import *
import openpyxl

window = Tk()
window.geometry('600x600')
window.title('Welcome to GYM fucking slave')

#Adding lbl
label = Label(window, text='Enter excel filename', font=('Arial Bold', 20))
label.grid(column=0, row=0)

label = Label(window, text=' Enter product code(s)', font=('Arial Bold', 20))
label.grid(column=0, row=1)


#Adding txtField
txt = Entry(window, width=40)
txt.grid(column=1,row=0)

txt_2 = Entry(window, width=40)
txt_2.grid(column=1,row=1)

#Adding open_file btn logic
def open_file():
    filetypes = (
        ('xlsx files', '*.xlsx'),
        ('po files', '*.po'),
        ('All files', '*.*')
    )
    window.update()
    file = filedialog.askopenfilename(filetypes=filetypes)
    name = os.path.basename(f'{file}')
    txt.insert(0, name)
    print(name)

#Adding button
btn_1 = Button(window, text='Open file', command=open_file)
btn_1.grid(column=0,row=8)

#Adding btn logic
def clicked():
    #Enter xlsx filname
    xlsx_name = '{}'.format(txt.get())
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

    messagebox.showinfo('Zip generated', 'Done')
    txt.delete(0, END)
    txt.insert(0, '')
    txt_2.delete(0, END)
    txt_2.insert(0, '')


#Adding button
btn = Button(window, text='Generate zip', command=clicked)
btn.grid(column=1, row=8)

window.mainloop()
