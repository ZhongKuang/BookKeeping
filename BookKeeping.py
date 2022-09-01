'''
查询当月总消费与总收入，需提前选好月份
记录消费与收入时，点击记录数据以保存
运行程序时不要打开文件，会同时占用
'''

from fileinput import filename
from glob import glob
from tkinter.simpledialog import *
from tkinter import *
import tkinter
from tkinter import *
from tkinter import messagebox
import tkinter
import tkinter.filedialog
import time
import os
import openpyxl
import datetime
from tkinter.ttk import *
import pandas as pd


'''
函数部分
save:保存已记录数据
gettime:获取实时时间
newbook:创建一个新的账本
new:根据filename创建账本
open:选择一个已有账本，如未选择，则默认回到初始账本
AllIncome:计算本月份总收入
AllConsume:计算本月份总消费
'''

def save():
        t_1 = t1.get()
        if t_1 == '' or not t_1.isdigit():
            messagebox.showinfo('提示', message='消费金额未输入或者消费金额不为数字')
        else:
            list = ['餐饮','交通','服装','日化','娱乐','收入']
            data = openpyxl.load_workbook(filename+'.xlsx')
            data1 = data.active
            a = datetime.datetime.now()
            b = list[comb1.current()]
            c = t1.get()
            data1.append([a, b, c])
            data1.column_dimensions['A'].width = 20
            data.save(filename+'.xlsx')
            messagebox.showinfo('提示', message='数据记录完毕')

def gettime():
      timestr = time.strftime("%H:%M:%S") # 获取当前的时间并转化为字符串
      lb.configure(text='time: '+timestr)   # 重新设置标签文本
      root.after(1000,gettime) # 每隔1s调用函数 gettime 自身获取时间

def newbook(file):
        global root
        data = openpyxl.Workbook()
        data1 = data.active
        data1['A1'] = '日期'
        data1['B1'] = '种类'
        data1['C1'] = '金额'
        data.save(file+'.xlsx')
        root.title(file+'.xlsx')

def new():
    global filename
    global root
    file=askstring('请输入','请输入新帐本名字')
    filename=file
    newbook(filename)
    root.title(filename+'.xlsx')

def open():
    global filename
    global filepath
    global root
    filepath=tkinter.filedialog.askopenfilename()
    filepath, filename = os.path.split(filepath)
    filename = filename.split('.')[0]
    if filename == '':
        if not os.path.exists(filename+'.xlsx'):
            newbook('AccountBook')
        filename='AccountBook'
        filepath=os.getcwd()
    root.title(filename+'.xlsx')

def AllIncome():
    data = pd.read_excel(filepath+'/'+filename+'.xlsx')
    data.dropna()
    data['月份'] = [str(i).split('T')[0].split('-')[1] for i in data['日期']]
    data['年份'] = [str(i).split('T')[0].split('-')[0] for i in data['日期']]   
    list = ['01','02','03','04','05','06','07','08','09','10','11','12']
    month = list[comb2.current()]
    data_Income = data[data['种类']=='收入']
    data_Income = data_Income[data_Income['月份']==month]
    if data_Income.empty:
        lb_3.config(text='0')
        messagebox.showinfo('提示', message='该月无记录')

    else:
        b = data_Income['金额'].sum()
        lb_3.config(text=b)

def AllConsume():
    data = pd.read_excel(filepath+'/'+filename+'.xlsx')
    data.dropna()
    data['月份'] = [str(i).split('T')[0].split('-')[1] for i in data['日期']]
    data['年份'] = [str(i).split('T')[0].split('-')[0] for i in data['日期']]   
    list = ['01','02','03','04','05','06','07','08','09','10','11','12']
    month = list[comb2.current()]
    data_consume = data[data['种类']!='收入']
    data_consume = data_consume[data_consume['月份']==month]
    if data_consume.empty:
        lb_4.config(text='0')
        messagebox.showinfo('提示', message='该月无记录')

    else:
        b = data_consume['金额'].sum()
        lb_4.config(text=b)
        

if __name__ == "__main__":
    '''
    
    '''
    filename=''
    filepath=os.getcwd()

    '''
    打开文件,初始化窗口
    '''
    root = Tk()
    open()
    root.geometry('240x240')

    '''
    记录数据
    '''
    #添加组合框
    var1 = StringVar()

    lb1 = Label(root, text='种类:', font=('微软雅黑', 10))
    lb1.grid(row=1,column=0)
    comb1 = Combobox(root,textvariable=var1,values=['餐饮','交通','服装','日化','娱乐','收入'])
    comb1.grid(row=1,column=1)

    lb2 = Label(root, text='金额:', font=('微软雅黑', 10))
    lb2.grid(row=2,column=0)
    t1 = Entry(root, font=('微软雅黑', 12), width=8)
    t1.grid(row=2, column=1)
    lb_1 = Label(root, text='元', font=('微软雅黑', 10))
    lb_1.grid(row=2,column=2)

    var2 = StringVar()

    lb3 = Label(root, text='月份:', font=('微软雅黑', 10))
    lb3.grid(row=6,column=0)
    comb2 = Combobox(root,textvariable=var2,values=['01','02','03','04','05','06','07','08','09','10','11','12'])
    comb2.grid(row=6,column=1)



    '''
    时间模块
    '''
    lb = tkinter.Label(root,text='',fg='black',font=("黑体",10))
    lb.grid(row=10,column=0)
    gettime()
    root.geometry('480x240')


    '''
    按钮模块
    '''
    Button1 = Button(root, text='记录数据', width=8, command=save)
    Button1.grid(row=7, column=0, sticky=W)
    Button2 = Button(root, text='退出',  width=8, command=root.quit)
    Button2.grid(row=7, column=1, sticky=E)
    Button3 = Button(root, text='当月总收入:', width=10, command=AllIncome)
    Button3.grid(row=4, column=0, sticky=W)
    Button4 = Button(root, text='当月总消费:',  width=10, command=AllConsume)
    Button4.grid(row=5, column=0, sticky=W)
    lb_3 = Label(root)
    lb_3.grid(row=4,column=1)
    lb_4 = Label(root)
    lb_4.grid(row=5,column=1)


    '''
    菜单分组 menuFile
    '''
    mainmenu = Menu(root)
    menuFile = Menu(mainmenu)  
    mainmenu.add_cascade(label="文件",menu=menuFile)
    menuFile.add_command(label="创建新账本",command=new)
    menuFile.add_command(label="选择账本",command=open)
    root.config(menu=mainmenu)

    root.mainloop()