import pandas as pd
import tkinter as tk
#welcome to Sanjay File Handling system

window = tk.Tk()


def get_textbox():
    global path
    path = textbox.get()


def get_exceldata():
    file = pd.read_excel(path)
    excel_file = pd.ExcelFile(path)
    sheets = excel_file.sheet_names
    for sheet in range(len(sheets)):
        if sheet == 0:
            df = excel_file.parse(sheets[sheet])
            continue
        if sheet < len(sheets):
            df = pd.merge(df, excel_file.parse(sheets[sheet]))

    ps_no = int(textbox2.get())
    df1 = df.set_index('ps number')
    newfdf = pd.Series(df1.loc[ps_no])
    file_name = str(ps_no) + '.xlsx'
    print(newfdf)
    newfdf.to_excel(file_name)


label = tk.Label(window, text='Enter path of the excel file')
label.pack()
textbox = tk.Entry()
textbox.pack()
button = tk.Button(window, text='search', command=get_textbox)
button.pack()
label2=tk.Label(window,text="enter the ps number you want to search")
label2.pack()
textbox2=tk.Entry()
textbox2.pack()
button2=tk.Button(window,text="get data",command=get_exceldata)
button2.pack()
window.mainloop()

