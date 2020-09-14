#!/usr/bin/env python
# -*- coding:utf-8 -*-
import process_quanbu_bumen
import process_zhenjie
import process_bumen
import process_quanbu_zhenjie
# import sys
# if sys.version_info[0] == 2:
#     from tkinter import *
#
#
#     #Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')
#     #import tkFileDialog
#     #import tkSimpleDialog
# else:  #Python 3.x
from tkinter import *
from tkinter.font import Font
from tkinter.ttk import *
from tkinter.messagebox import *
from tkinter import filedialog


# import tkinter.filedialog as tkFileDialog
# import tkinter.simpledialog as tkSimpleDialog    #askstring()

class Application_ui(Frame):
    # 这个类仅实现界面生成功能，具体事件处理代码在子类Application中。
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('学习强国数据统计助手-宣传部-3.0')
        self.master.geometry('724x584')
        self.createWidgets()

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.style.configure('TLabel1.TLabel', anchor='w', font=('宋体', 9))
        self.Label1 = Label(self.top, text='请选择对应表', style='TLabel1.TLabel')
        self.Label1.place(relx=0.144, rely=0.151, relwidth=0.178, relheight=0.029)
        self.text1 = Entry(self.top, width=95)
        self.text1.place(relx=0.344, rely=0.151, relwidth=0.278, relheight=0.029)
        self.button1 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile)
        self.button1.place(relx=0.644, rely=0.151, relwidth=0.178, relheight=0.039)
        # self.text1.insert(INSERT, '/Users/wuyiquan/PycharmProjects/tongjizhushou/单位对应表.xlsx')

        self.style.configure('TLabel8.TLabel', anchor='w', font=('宋体', 9))
        self.Label8 = Label(self.top, text='请选择顺序', style='TLabel8.TLabel')
        self.Label8.place(relx=0.144, rely=0.200, relwidth=0.178, relheight=0.029)
        self.text8 = Entry(self.top, width=95)
        self.text8.place(relx=0.344, rely=0.200, relwidth=0.278, relheight=0.029)
        self.button8 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile8)
        self.button8.place(relx=0.644, rely=0.200, relwidth=0.178, relheight=0.039)
        # self.text8.insert(INSERT, '/Users/wuyiquan/PycharmProjects/tongjizhushou/单位顺序.xlsx')

        self.style.configure('TLabel7.TLabel', anchor='w', font=('宋体', 9))
        self.Label7 = Label(self.top, text='请选择镇街对应表', style='TLabel7.TLabel')
        self.Label7.place(relx=0.144, rely=0.250, relwidth=0.178, relheight=0.029)
        self.text7 = Entry(self.top, width=95)
        self.text7.place(relx=0.344, rely=0.250, relwidth=0.278, relheight=0.029)
        self.button7 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile7)
        self.button7.place(relx=0.644, rely=0.250, relwidth=0.178, relheight=0.039)

        self.style.configure('TLabel9.TLabel', anchor='w', font=('宋体', 9))
        self.Label9 = Label(self.top, text='请选择镇街顺序', style='TLabel9.TLabel')
        self.Label9.place(relx=0.144, rely=0.300, relwidth=0.178, relheight=0.029)
        self.text9 = Entry(self.top, width=95)
        self.text9.place(relx=0.344, rely=0.300, relwidth=0.278, relheight=0.029)
        self.button9 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile9)
        self.button9.place(relx=0.644, rely=0.300, relwidth=0.178, relheight=0.039)

        self.style.configure('TLabel2.TLabel', anchor='w', font=('宋体', 9))
        self.Label2 = Label(self.top, text='请选择区直表', style='TLabel2.TLabel')
        self.Label2.place(relx=0.144, rely=0.354, relwidth=0.178, relheight=0.029)
        self.text2 = Entry(self.top, width=95)
        self.text2.place(relx=0.344, rely=0.354, relwidth=0.278, relheight=0.029)
        self.button2 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile2)
        self.button2.place(relx=0.644, rely=0.354, relwidth=0.178, relheight=0.039)
        # self.text2.insert(INSERT, '/Users/wuyiquan/PycharmProjects/tongjizhushou/0816数据/16日数据情况区直.xlsx')

        self.style.configure('TLabel3.TLabel', anchor='w', font=('宋体', 9))
        self.Label3 = Label(self.top, text='请选择党内表', style='TLabel3.TLabel')
        self.Label3.place(relx=0.144, rely=0.400, relwidth=0.178, relheight=0.029)
        self.text3 = Entry(self.top, width=95)
        self.text3.place(relx=0.344, rely=0.400, relwidth=0.278, relheight=0.029)
        self.button3 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile3)
        self.button3.place(relx=0.644, rely=0.400, relwidth=0.178, relheight=0.039)
        # self.text3.insert(INSERT, '/Users/wuyiquan/PycharmProjects/tongjizhushou/0816数据/中共杭州市西湖区委宣传部下级组织2020%2F08%2F16日数据情况党内.xlsx')

        self.style.configure('TLabel4.TLabel', anchor='w', font=('宋体', 9))
        self.Label4 = Label(self.top, text='请选择党外表', style='TLabel4.TLabel')
        self.Label4.place(relx=0.144, rely=0.466, relwidth=0.178, relheight=0.029)
        self.text4 = Entry(self.top, width=95)
        self.text4.place(relx=0.344, rely=0.466, relwidth=0.278, relheight=0.029)
        self.button4 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile4)
        self.button4.place(relx=0.644, rely=0.466, relwidth=0.178, relheight=0.039)
        # self.text4.insert(INSERT, '/Users/wuyiquan/PycharmProjects/tongjizhushou/0816数据/西湖区学习组团下级组织2020_08_16日数据情况党外.xlsx')

        self.style.configure('TLabel5.TLabel', anchor='w', font=('宋体', 9))
        self.Label5 = Label(self.top, text='请选择目标文件夹', style='TLabel5.TLabel')
        self.Label5.place(relx=0.144, rely=0.562, relwidth=0.178, relheight=0.029)
        self.text5 = Entry(self.top, width=95)
        self.text5.place(relx=0.344, rely=0.562, relwidth=0.278, relheight=0.029)
        self.button5 = Button(self.top, text='浏览', width=2, command=self.selectExcelfile5)
        self.button5.place(relx=0.644, rely=0.562, relwidth=0.178, relheight=0.039)
        # self.text5.insert(INSERT, '/Users/wuyiquan/Downloads')

        self.style.configure('TCommand1.TButton', font=('宋体', 9))
        self.Command1 = Button(self.top, text='启动-全部-部门', command=self.Command1_Cmd, style='TCommand1.TButton')
        self.Command1.place(relx=0.109, rely=0.671, relwidth=0.156, relheight=0.07)

        self.style.configure('TCommand4.TButton', font=('宋体', 9))
        self.Command4 = Button(self.top, text='启动-全部-镇街', command=self.Command4_Cmd, style='TCommand4.TButton')
        self.Command4.place(relx=0.309, rely=0.671, relwidth=0.156, relheight=0.07)

        self.style.configure('TCommand2.TButton', font=('宋体', 9))
        self.Command2 = Button(self.top, text='启动-通报-镇街', command=self.Command2_Cmd, style='TCommand2.TButton')
        self.Command2.place(relx=0.509, rely=0.671, relwidth=0.156, relheight=0.07)

        self.style.configure('TCommand3.TButton', font=('宋体', 9))
        self.Command3 = Button(self.top, text='启动-通报-部门', command=self.Command3_Cmd, style='TCommand3.TButton')
        self.Command3.place(relx=0.709, rely=0.671, relwidth=0.156, relheight=0.07)



        self.Text1Var = StringVar(value='您好，欢迎使用本工具。开发者：吴亦全')
        self.Text1 = Entry(self.top, textvariable=self.Text1Var, font=('宋体', 9))
        self.Text1.place(relx=0.144, rely=0.795, relwidth=0.698, relheight=0.125)

        self.style.configure('TLabel6.TLabel', anchor='w', font=('文鼎CS大黑', 22))
        self.Label6 = Label(self.top, text='学习强国数据统计助手', style='TLabel6.TLabel')
        self.Label6.place(relx=0.298, rely=0.027, relwidth=0.521, relheight=0.084)


class Application(Application_ui):

    # 这个类实现具体的事件处理回调函数。界面生成代码在Application_ui中。
    def __init__(self, master=None):
        Application_ui.__init__(self, master)

    def Command1_Cmd(self, event=None):
        # TODO, Please finish the function here!
        # 在这里审核不要出问题
        self.Text1.insert(END, "正在自动计算中……")
        process_quanbu_bumen.work_package(self.text1.get(), self.text8.get(), self.text2.get(), self.text3.get(),
                                    self.text4.get(), self.text5.get(), self.Text1)

    def Command2_Cmd(self, event=None):
        # TODO, Please finish the function here!
        # 在这里审核不要出问题
        self.Text1.insert(END, "正在自动计算中……")
        process_zhenjie.work_package(self.text7.get(), self.text9.get(), self.text3.get(), self.text4.get(),
                                     self.text5.get(), self.Text1)

    def Command3_Cmd(self, event=None):
        # TODO, Please finish the function here!
        # 在这里审核不要出问题
        self.Text1.insert(END, "正在自动计算中……")
        process_bumen.work_package(self.text1.get(), self.text8.get(), self.text2.get(), self.text3.get(), self.text4.get(),
                                   self.text5.get(), self.Text1)

    def Command4_Cmd(self, event=None):
        # TODO, Please finish the function here!
        # 在这里审核不要出问题
        self.Text1.insert(END, "正在自动计算中……")
        process_quanbu_zhenjie.work_package(self.text7.get(), self.text9.get(), self.text3.get(), self.text4.get(),
                                     self.text5.get(), self.Text1)

    # 先清空
    def selectExcelfile(self):
        self.text1.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel2', '*.xlsx')])
        print(sfname)
        self.text1.insert(INSERT, sfname)

    def selectExcelfile8(self):
        self.text8.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel8', '*.xlsx')])
        print(sfname)
        self.text8.insert(INSERT, sfname)

    def selectExcelfile7(self):
        self.text7.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel7', '*.xlsx')])
        print(sfname)
        self.text7.insert(INSERT, sfname)

    def selectExcelfile9(self):
        self.text9.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel9', '*.xlsx')])
        print(sfname)
        self.text9.insert(INSERT, sfname)

    def selectExcelfile2(self):
        self.text2.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel2', '*.xlsx')])
        print(sfname)
        self.text2.insert(INSERT, sfname)

    def selectExcelfile3(self):
        self.text3.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel2', '*.xlsx')])
        print(sfname)
        self.text3.insert(INSERT, sfname)

    def selectExcelfile4(self):
        self.text4.delete(0, END)
        sfname = filedialog.askopenfilename(title='选择Excel文件', filetypes=[('Excel2', '*.xlsx')])
        print(sfname)
        self.text4.insert(INSERT, sfname)

    def selectExcelfile5(self):
        self.text5.delete(0, END)
        sfname = filedialog.askdirectory(title='选择要保存的文件夹')
        print(sfname)
        self.text5.insert(INSERT, sfname)


if __name__ == "__main__":
    top = Tk()
    Application(top).mainloop()
