#! /usr/bin/env python
#  -*- coding: utf-8 -*-
#
# Jun 15, 2021 12:00:53 AM CST  platform: Windows NT
# author:halfmoon

import sys

import tkinter as tk
from tkinter.constants import END
import tkinter.messagebox
import tkinter.filedialog
import tkinter.ttk as ttk
import pandas as pd
from tksheet import Sheet
import time
from tabulate import tabulate

import data
import gui_support

# 文件读取
filename = ''
datForSheet = []
dat = []

# 数据属性
outmin = ''
outmax = ''
outava = ''
outmost = ''
outmid = ''

# 写报告
contain = '''# 数据分析报告\n'''
contain += '由SDAT于'+str(time.asctime())+'生成\n'
sheetContain = ''
analyContain = ''
img1Contain = '![Plot](./asset/Plot.png)\n'
img2Contain = '![Bar](./asset/Bar.png)\n'

# load将文件名赋给filename并读取数据
def load_Excel():
    global datForSheet
    global dat
    global filename
    print('load_Execl')
    sys.stdout.flush()
    filename = tk.filedialog.askopenfilename(filetypes=[("xls", ".xls"),("xlsx", ".xlsx")])
    print(filename)
    dat = pd.read_excel(filename,   # filepath here
                                    #engine = "openpyxl",
                                    )# header = None
    datForSheet = dat.values.tolist()
    

def load_Csv():
    global datForSheet
    global dat
    global filename
    print('load_Csv')
    sys.stdout.flush()
    filename =  tk.filedialog.askopenfilename(filetypes=[("CSV",".csv")])
    print(filename)
    dat = pd.read_csv(filename,   # filepath here
                        encoding='gbk')# 防止乱码，应该是pandas不能准确识别
    datForSheet = dat.values.tolist()

def load_Txt():
    global datForSheet
    global dat
    global filename
    print('load_Txt')
    sys.stdout.flush()
    filename =  tk.filedialog.askopenfilename(filetypes=[('TXT',".txt")])
    print(filename)
    dat = pd.read_table(filename,   # filepath here
                                    )# header = None
    datForSheet = dat.values.tolist()


def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = tk.Tk()
    top = Main (root)
    gui_support.init(root, top)
    root.mainloop()

w = None
def create_Main(rt, *args, **kwargs):
    '''Starting point when module is imported by another module.
       Correct form of call: 'create_Main(root, *args, **kwargs)' .'''
    global w, w_win, root
    #rt = root
    root = rt
    w = tk.Toplevel (root)
    top = Main (w)
    gui_support.init(w, top, *args, **kwargs)
    return (w, top)

def destroy_Main():
    global w
    w.destroy()
    w = None

class Main:
    def __init__(self, top=None):
        global sheetContain
        global analyContain
        global img1Contain
        global img2Contain
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#ffffff'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#ececec' # Closest X11 color: 'gray92'
        self.style = ttk.Style()
        if sys.platform == "win32":
            self.style.theme_use('winnative')
        self.style.configure('.',background=_bgcolor)
        self.style.configure('.',foreground=_fgcolor)
        self.style.configure('.',font="TkDefaultFont")
        self.style.map('.',background=
            [('selected', _compcolor), ('active',_ana2color)])

        top.geometry("803x605+332+104")
        top.minsize(120, 120)
        top.maxsize(1540, 825)
        top.resizable(1,  1)
        top.title("SDAT - A Simple Data Analysis Tool")
        top.configure(background="#ffffff")
        top.configure(highlightbackground="#ffffff")
        top.configure(highlightcolor="black")

        self.menubar = tk.Menu(top,font="TkMenuFont",bg=_bgcolor,fg=_fgcolor)
        top.configure(menu = self.menubar)

        self.sub_menu = tk.Menu(top,
                activebackground="#ececec",
                activeborderwidth=1,
                activeforeground="#000000",
                background="#ffffff",
                borderwidth=1,
                disabledforeground="#a3a3a3",
                foreground="#000000",
                tearoff=0)
        self.menubar.add_cascade(menu=self.sub_menu,
                label="加载数据...")
        self.sub_menu.add_command(
                command=lambda:(load_Excel(),self.Scrolledtext1.insert(END,'已加载'+filename+'\n'+'请点击开始分析\n'),\
                    self.Scrolledtext1.see(END)),
                label="Excel文件")
        self.sub_menu.add_command(
                command=lambda:(load_Csv(),self.Scrolledtext1.insert(END,'已加载'+filename+'\n'+'请点击开始分析\n'),\
                    self.Scrolledtext1.see(END)),
                label="csv文件")
        self.sub_menu.add_command(
                command=lambda:load_Txt(),
                label="txt文件")
        self.menubar.add_command(
                command=lambda:self.refreshSheet(),
                label="开始分析")
        self.sub_menu1 = tk.Menu(top,
                activebackground="#ececec",
                activeborderwidth=1,
                activeforeground="#000000",
                background="#ffffff",
                borderwidth=1,
                disabledforeground="#a3a3a3",
                foreground="#000000",
                tearoff=0)
        self.menubar.add_cascade(menu=self.sub_menu1,
                label="导出报告...")
        self.sub_menu1.add_command(
                command=lambda:[gui_support.saveMd(contain,sheetContain,analyContain,img1Contain,img2Contain),\
                    self.Scrolledtext1.insert(END,'Report.md已保存在./ 中\n'),self.Scrolledtext1.see(END)],
                label="markdown格式")
        self.menubar.add_command(
                command=lambda:tk.messagebox.showinfo('About',
                '''Halfmoon \n i.e 詹骋昊'''),
                label="关于")

        self.TPanedwindow1 = ttk.Panedwindow(top, orient="horizontal")
        self.TPanedwindow1.place(relx=0.0, rely=0.0, relheight=1.0, relwidth=1.0)

        self.TPanedwindow1_p1 = ttk.Labelframe(self.TPanedwindow1, width=75
                , text='数据表格')
        self.TPanedwindow1.add(self.TPanedwindow1_p1, weight=0)

        self.TPanedwindow1_p2 = ttk.Labelframe(self.TPanedwindow1, text='信息与选项')
        self.TPanedwindow1.add(self.TPanedwindow1_p2, weight=0)
        self.__funcid0 = self.TPanedwindow1.bind('<Map>', self.__adjust_sash0)

        self.TButton1 = ttk.Button(self.TPanedwindow1_p1)
        self.TButton1.place(relx=0.667, rely=0.0, height=27, width=87
                 , bordermode='ignore')
        self.TButton1.configure(takefocus="")
        self.TButton1.configure(text='''开始分析！''')
        self.TButton1.configure(command=lambda:self.refreshSheet())


# 右半板块
        
        self.TLabelframe1 = ttk.Labelframe(self.TPanedwindow1_p2)
        self.TLabelframe1.place(relx=0.024, rely=0.06, relheight=0.27
                , relwidth=0.96, bordermode='ignore')
        self.TLabelframe1.configure(relief='')
        self.TLabelframe1.configure(text='''常见描述统计量''')

        self.TLabel1 = ttk.Label(self.TLabelframe1)
        self.TLabel1.place(relx=0.025, rely=0.32, height=22, width=120
                , bordermode='ignore')
        self.TLabel1.configure(background="#ffffff")
        self.TLabel1.configure(foreground="#000000")
        self.TLabel1.configure(font="TkDefaultFont")
        self.TLabel1.configure(relief="flat")
        self.TLabel1.configure(anchor='w')
        self.TLabel1.configure(justify='left')
        self.TLabel1.configure(text='''平均数：'''+outava)

        self.TLabel2 = ttk.Label(self.TLabelframe1)
        self.TLabel2.place(relx=0.315, rely=0.32, height=22, width=120
                , bordermode='ignore')
        self.TLabel2.configure(background="#ffffff")
        self.TLabel2.configure(foreground="#000000")
        self.TLabel2.configure(font="TkDefaultFont")
        self.TLabel2.configure(relief="flat")
        self.TLabel2.configure(anchor='w')
        self.TLabel2.configure(justify='left')
        self.TLabel2.configure(text='''中位数：'''+outmid)

        self.TLabel3 = ttk.Label(self.TLabelframe1)
        self.TLabel3.place(relx=0.605, rely=0.32, height=22, width=120
                , bordermode='ignore')
        self.TLabel3.configure(background="#ffffff")
        self.TLabel3.configure(foreground="#000000")
        self.TLabel3.configure(font="TkDefaultFont")
        self.TLabel3.configure(relief="flat")
        self.TLabel3.configure(anchor='w')
        self.TLabel3.configure(justify='left')
        self.TLabel3.configure(text='''众数：'''+outmost)

        self.TLabel4 = ttk.Label(self.TLabelframe1)
        self.TLabel4.place(relx=0.025, rely=0.62, height=22, width=120
                , bordermode='ignore')
        self.TLabel4.configure(background="#ffffff")
        self.TLabel4.configure(foreground="#000000")
        self.TLabel4.configure(font="TkDefaultFont")
        self.TLabel4.configure(relief="flat")
        self.TLabel4.configure(anchor='w')
        self.TLabel4.configure(justify='left')
        self.TLabel4.configure(text='''最小值：'''+outmin)

        self.TLabel5 = ttk.Label(self.TLabelframe1)
        self.TLabel5.place(relx=0.315, rely=0.62, height=22, width=120
                , bordermode='ignore')
        self.TLabel5.configure(background="#ffffff")
        self.TLabel5.configure(foreground="#000000")
        self.TLabel5.configure(font="TkDefaultFont")
        self.TLabel5.configure(relief="flat")
        self.TLabel5.configure(anchor='w')
        self.TLabel5.configure(justify='left')
        self.TLabel5.configure(text='''最大值：'''+outmax)

        self.TLabelframe2 = ttk.Labelframe(self.TPanedwindow1_p2)
        self.TLabelframe2.place(relx=0.025, rely=0.35, relheight=0.3
                , relwidth=0.96, bordermode='ignore')
        self.TLabelframe2.configure(relief='')
        self.TLabelframe2.configure(text='''图例''')

        self.TButton2 = ttk.Button(self.TLabelframe2)
        self.TButton2.place(relx=0.075, rely=0.6, height=27, width=87
                , bordermode='ignore')
        self.TButton2.configure(takefocus="")
        self.TButton2.configure(text='''生成并预览''')
        self.TButton2.configure(command=lambda:[data.savePlot(dat),\
            self.Scrolledtext1.insert(END,'图片已保存在./asset \n'),self.Scrolledtext1.see(END)])


        self.TButton3 = ttk.Button(self.TLabelframe2)
        self.TButton3.place(relx=0.4, rely=0.6, height=27, width=87
                , bordermode='ignore')
        self.TButton3.configure(takefocus="")
        self.TButton3.configure(text='''生成并预览''')
        self.TButton3.configure(command=lambda:[data.saveBar(dat),\
            self.Scrolledtext1.insert(END,'图片已保存在./asset \n'),self.Scrolledtext1.see(END)])


        self.TLabel6 = ttk.Label(self.TLabelframe2)
        self.TLabel6.place(relx=0.125, rely=0.3, height=21, width=39
                , bordermode='ignore')
        self.TLabel6.configure(background="#ffffff")
        self.TLabel6.configure(foreground="#000000")
        self.TLabel6.configure(font="TkDefaultFont")
        self.TLabel6.configure(relief="flat")
        self.TLabel6.configure(anchor='w')
        self.TLabel6.configure(justify='left')
        self.TLabel6.configure(text='''折线图''')

        self.TLabel6 = ttk.Label(self.TLabelframe2)
        self.TLabel6.place(relx=0.45, rely=0.3, height=21, width=39
                , bordermode='ignore')
        self.TLabel6.configure(background="#ffffff")
        self.TLabel6.configure(foreground="#000000")
        self.TLabel6.configure(font="TkDefaultFont")
        self.TLabel6.configure(relief="flat")
        self.TLabel6.configure(anchor='w')
        self.TLabel6.configure(justify='left')
        self.TLabel6.configure(text='''直方图''')

        self.TLabelframe3 = ttk.Labelframe(self.TPanedwindow1_p2)
        self.TLabelframe3.place(relx=0.025, rely=0.67, relheight=0.30
                , relwidth=0.96, bordermode='ignore')
        self.TLabelframe3.configure(relief='')
        self.TLabelframe3.configure(text='''提示信息''')

        self.Scrolledtext1 = ScrolledText(self.TLabelframe3)
        self.Scrolledtext1.place(relx=0.027, rely=0.182, relheight=0.782
                , relwidth=0.959, bordermode='ignore')
        self.Scrolledtext1.configure(background="white")
        self.Scrolledtext1.configure(font="TkTextFont")
        self.Scrolledtext1.configure(foreground="black")
        self.Scrolledtext1.configure(highlightbackground="#d9d9d9")
        self.Scrolledtext1.configure(highlightcolor="black")
        self.Scrolledtext1.configure(insertbackground="black")
        self.Scrolledtext1.configure(insertborderwidth="3")
        self.Scrolledtext1.configure(selectbackground="blue")
        self.Scrolledtext1.configure(selectforeground="white")
        self.Scrolledtext1.configure(wrap="none")


# 刷新以更新表格（按钮及菜单“开始分析”）
    def refreshSheet(self):
        global outava
        global outmax
        global outmin
        global outmost
        global outmid
        global sheetContain
        global analyContain
        print('Refresh!')
        if datForSheet != []:
            self.sheet = Sheet(self.TPanedwindow1_p1,data=datForSheet)
            self.sheet.enable_bindings()
            self.sheet.place(relheight=0.94,relwidth=0.96,relx=0.025,y=25)
            outava = str(dat[dat.columns[1]].mean())
            outmid = str(dat[dat.columns[1]].median())
            outmax = str(dat[dat.columns[1]].max())
            outmin = str(dat[dat.columns[1]].min())
            outmost = str(dat[dat.columns[1]].mode()[0])
            sheetContain ='## 数据表格\n'\
                +tabulate(dat,tablefmt="pipe",headers=dat.columns)+'\n'
            analyContain ='## 单变量分析\n'+'- 平均数：'+outava+'\n'+\
                '- 中位数：'+outmid+'\n'+'- 众数：'+outmost+'\n'\
                +'- 最小值：'+outmin+'\n'\
                +'- 最大值：'+outmax+'\n'
            if len(dat.columns) > 2:
                self.Scrolledtext1.insert(END,'目前仅支持单变量分析，导入数据存在异常，请检查\n')
            else:
                self.Scrolledtext1.insert(END,'分析完成\n')
            self.Scrolledtext1.see(END)

        self.TLabel1.configure(text='''平均数：'''+outava)
        self.TLabel2.configure(text='''中位数：'''+outmid)
        self.TLabel3.configure(text='''众数：'''+outmost)
        self.TLabel4.configure(text='''最小值：'''+outmin)
        self.TLabel5.configure(text='''最大值：'''+outmax)
        self.sheet.update()

        
# 自动生成
    def __adjust_sash0(self, event):
        paned = event.widget
        pos = [285, ]
        i = 0
        for sash in pos:
            paned.sashpos(i, sash)
            i += 1
        paned.unbind('<map>', self.__funcid0)
        del self.__funcid0

# The following code is added to facilitate the Scrolled widgets you specified.
class AutoScroll(object):
    '''Configure the scrollbars for a widget.'''
    def __init__(self, master):
        #  Rozen. Added the try-except clauses so that this class
        #  could be used for scrolled entry widget for which vertical
        #  scrolling is not supported. 5/7/14.
        try:
            vsb = ttk.Scrollbar(master, orient='vertical', command=self.yview)
        except:
            pass
        hsb = ttk.Scrollbar(master, orient='horizontal', command=self.xview)
        try:
            self.configure(yscrollcommand=self._autoscroll(vsb))
        except:
            pass
        self.configure(xscrollcommand=self._autoscroll(hsb))
        self.grid(column=0, row=0, sticky='nsew')
        try:
            vsb.grid(column=1, row=0, sticky='ns')
        except:
            pass
        hsb.grid(column=0, row=1, sticky='ew')
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(0, weight=1)
        # Copy geometry methods of master  (taken from ScrolledText.py)
        methods = tk.Pack.__dict__.keys() | tk.Grid.__dict__.keys() \
                  | tk.Place.__dict__.keys()
        for meth in methods:
            if meth[0] != '_' and meth not in ('config', 'configure'):
                setattr(self, meth, getattr(master, meth))

    @staticmethod
    def _autoscroll(sbar):
        '''Hide and show scrollbar as needed.'''
        def wrapped(first, last):
            first, last = float(first), float(last)
            if first <= 0 and last >= 1:
                sbar.grid_remove()
            else:
                sbar.grid()
            sbar.set(first, last)
        return wrapped

    def __str__(self):
        return str(self.master)

def _create_container(func):
    '''Creates a ttk Frame with a given master, and use this new frame to
    place the scrollbars and the widget.'''
    def wrapped(cls, master, **kw):
        container = ttk.Frame(master)
        container.bind('<Enter>', lambda e: _bound_to_mousewheel(e, container))
        container.bind('<Leave>', lambda e: _unbound_to_mousewheel(e, container))
        return func(cls, container, **kw)
    return wrapped

class ScrolledText(AutoScroll, tk.Text):
    '''A standard Tkinter Text widget with scrollbars that will
    automatically show/hide as needed.'''
    @_create_container
    def __init__(self, master, **kw):
        tk.Text.__init__(self, master, **kw)
        AutoScroll.__init__(self, master)

import platform
def _bound_to_mousewheel(event, widget):
    child = widget.winfo_children()[0]
    if platform.system() == 'Windows' or platform.system() == 'Darwin':
        child.bind_all('<MouseWheel>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Shift-MouseWheel>', lambda e: _on_shiftmouse(e, child))
    else:
        child.bind_all('<Button-4>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Button-5>', lambda e: _on_mousewheel(e, child))
        child.bind_all('<Shift-Button-4>', lambda e: _on_shiftmouse(e, child))
        child.bind_all('<Shift-Button-5>', lambda e: _on_shiftmouse(e, child))

def _unbound_to_mousewheel(event, widget):
    if platform.system() == 'Windows' or platform.system() == 'Darwin':
        widget.unbind_all('<MouseWheel>')
        widget.unbind_all('<Shift-MouseWheel>')
    else:
        widget.unbind_all('<Button-4>')
        widget.unbind_all('<Button-5>')
        widget.unbind_all('<Shift-Button-4>')
        widget.unbind_all('<Shift-Button-5>')

def _on_mousewheel(event, widget):
    if platform.system() == 'Windows':
        widget.yview_scroll(-1*int(event.delta/120),'units')
    elif platform.system() == 'Darwin':
        widget.yview_scroll(-1*int(event.delta),'units')
    else:
        if event.num == 4:
            widget.yview_scroll(-1, 'units')
        elif event.num == 5:
            widget.yview_scroll(1, 'units')

def _on_shiftmouse(event, widget):
    if platform.system() == 'Windows':
        widget.xview_scroll(-1*int(event.delta/120), 'units')
    elif platform.system() == 'Darwin':
        widget.xview_scroll(-1*int(event.delta), 'units')
    else:
        if event.num == 4:
            widget.xview_scroll(-1, 'units')
        elif event.num == 5:
            widget.xview_scroll(1, 'units')

# 支持函数到此结束

if __name__ == '__main__':
    vp_start_gui()

