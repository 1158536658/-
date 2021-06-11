import os
import tkinter
import tkinter.messagebox
from tkinter import *
from tkinter.filedialog import *
import datetime
from tkinter.ttk import *
import pandas as pd
import threading



def formatForm(form, width, heigth):
    """设置居中显示"""
    # 得到屏幕宽度
    win_width = form.winfo_screenwidth()
    # 得到屏幕高度
    win_higth = form.winfo_screenheight()

    # 计算偏移量
    width_adjust = (win_width - width) / 2
    higth_adjust = (win_higth - heigth) / 2

    form.geometry("%dx%d+%d+%d" % (width, heigth, width_adjust, higth_adjust))

root = Tk()
root.title("Perfect Compare Tools for Excel V1.0")
# root['bg'] = '#DCDCDC'
# root.geometry('850x500')
# 创建窗口大小
formatForm(root, 850, 500)

# 声明变量
scale = 100
curr_time = datetime.datetime.now()
localtime = str(datetime.datetime.strftime(curr_time, '%Y%m%d%H%M%S'))
p_loc = localtime[:13]
path = os.getcwd() + '\\' + p_loc + '_compare.xlsx'


#
def select_compare_file():
    global name
    name = askopenfilename()
    file1_text.set(name)




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
    global path
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
                    # 设置单元格为空格
                    # dfDiff.iloc[row, col] = np.NAN
                else:
                    if value_OLD == 'nan' or value_NEW == 'nan':
                        dfDiff.iloc[row, col] = f2.iloc[row, col]

                    else:

                        dfDiff.iloc[row, col] = ('{}→{}').format(value_OLD, value_NEW)
            except:

                dfDiff.iloc[row, col] = ('{}→{}').format(value_OLD, 'NaN')


    # Save output and format
    try:
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        dfDiff.to_excel(writer, sheet_name='DIFF', index=False)
        # get xlsxwriter objects
        workbook = writer.book
        worksheet = writer.sheets['DIFF']
        # 隐藏excel的线条
        # worksheet.hide_gridlines(5)

        # define formats
        highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color': '#FFFF00'})

        # set format over range
        ## highlight changed cells
        worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                                   'criteria': 'containing',
                                                   'value': '→',
                                                   'format': highlight_fmt})
        # save
        writer.save()
        print('Compare Done.')
    except:
        tkinter.messagebox.showwarning(title='Hi', message='请关闭xlsx文件后重新进行操作')



def on_click():
    global path
    f1 = file1_text.get()
    f2 = file2_text.get()
    b3.configure(text="处理中,请不要操作...", state=DISABLED)
    try:
        compare_excel(f1, f2)

    except:
        tkinter.messagebox.showwarning(title='Hi', message='请先选择文件')
        b3.configure(text="Start Compare", state=NORMAL)
    else:
        tkinter.messagebox.showinfo(title='success', message="Compare done, please see the result: " + path)
        res()
        b3.configure(text="Start Compare", state=NORMAL)



def res():
    l_res = Label(root, text=path)
    l_res.pack()
    l_res.place(x=190, y=280)
    open_btn = tkinter.Button(root, text="打开对比文件", command=open_click)
    open_btn.pack()
    open_btn.place(x=190, y=320)



def open_click():
    global path
    os.startfile(path)


def thread_it(func, *args):
    '''将函数打包进线程'''
    # 创建
    t = threading.Thread(target=func, args=args)
    # 守护
    t.setDaemon(True)
    # 启动
    t.start()

b3 = tkinter.Button(root, text="Start Compare", command = lambda :thread_it(on_click))
b3.pack()
b3.place(x=360, y=180)


root.mainloop()