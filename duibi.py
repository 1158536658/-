import tkinter
import tkinter.messagebox
from tkinter import *
from tkinter.filedialog import *
import pandas as pd

root = Tk()
root.title("Perfect Compare Tools for Excel")
# root['bg'] = '#DCDCDC'
# root.geometry('600x250')
root.geometry('850x500')

def select_compare_file():
    name = askopenfilename()
    file1_text.set(name)

# Button(root, text="Select Compare File", command=select_compare_file).pack()
b1 = tkinter.Button(root, text="Select Compare File", command=select_compare_file)
b1.pack()
b1.place(x=530, y=60)

file1_text = StringVar()
file1 = Entry(root, textvariable=file1_text, width=65)
file1_text.set(" ")
file1.pack()
file1.place(x=60, y=60, height=29)
Label(root, text="").pack()




def select_to_file():
    a = askopenfilename()
    file2_text.set(a)
# Button(root, text="Select To File", command = select_to_file).pack()
b2 = tkinter.Button(root, text="Select To File", command = select_to_file)
b2.pack()
b2.place(x=530, y=100)
file2_text = StringVar()
file2 = Entry(root, textvariable=file2_text, width=65)
file2_text.set(" ")
file2.pack()
file2.place(x=60, y=100, height=29)

Label(root, text="").pack()

def compare_excel(compare, to):
    f1 = pd.read_excel(compare, keep_default_na=False)
    f2 = pd.read_excel(to, keep_default_na=False)

    # Perform Diff
    dfDiff = f1.copy()
    for row in range(dfDiff.shape[0]):
        for col in range(dfDiff.shape[1]):
            value_OLD = f1.iloc[row, col]
            try:
                value_NEW = f2.iloc[row, col]
                if value_OLD == value_NEW:
                    dfDiff.iloc[row, col] = f2.iloc[row, col]
                else:
                    dfDiff.iloc[row, col] = ('{}→{}').format(value_OLD, value_NEW)
            except:
                dfDiff.iloc[row, col] = ('{}→{}').format(value_OLD, 'NaN')

    # Save output and format
    writer = pd.ExcelWriter('compare_result.xlsx', engine='xlsxwriter')

    dfDiff.to_excel(writer, sheet_name='DIFF', index=False)

    # get xlsxwriter objects
    workbook = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.hide_gridlines(2)

    # define formats
    highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color': '#6A93B0'})

    # set format over range
    ## highlight changed cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                               'criteria': 'containing',
                                               'value': '→',
                                               'format': highlight_fmt})
    # save
    writer.save()
    print('Compare Done.')


def on_click():
    f1 = file1_text.get()
    f2 = file2_text.get()

    compare_excel(f1, f2)

    tkinter.messagebox.showinfo(title='success', message="Compare done, please see the result: 'compare_result.xlsx")
    # Label(root, text="Compare done, please see the result: 'compare_result.xlsx'").pack()
    # l1 = tkinter.Button(root, text="Compare done, please see the result: 'compare_result.xlsx'", fg='green' )
    # l1.pack()
    # l1.place(x=190, y=240)


# Button(root, text="Start Compare", command=on_click).pack()
b3 = tkinter.Button(root, text="Start Compare", command=on_click)
b3.pack()
b3.place(x=360, y=180)
root.mainloop()