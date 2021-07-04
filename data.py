# 绘图有关函数


import numpy
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import time

def savePlot(data):
    d = data[data.columns[1]]
    d.index = data[data.columns[0]]
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来正常显示中文标签
    plt.rcParams['axes.unicode_minus'] = False           # 用来正常显示负号
    fig = plt.figure(figsize=(8.0,5.6))
    plt.xlabel(data.columns[0])
    plt.ylabel(data.columns[1])
    d.plot()                            # 这直接用 pandas 封装的绘图
    plt.savefig('./asset/Plot.png')
    plt.show()

def saveBar(data):
    y = data[data.columns[1]]
    x = data[data.columns[0]]
    if str(type(x[0])) != '''<class 'str'>''':
        x = x.apply(lambda i: i.strftime('%Y-%m'))
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来正常显示中文标签
    plt.rcParams['axes.unicode_minus'] = False           # 用来正常显示负号
    fig = plt.figure(figsize=(8.0,5.6))
    plt.xlabel(data.columns[0])
    plt.ylabel(data.columns[1])
    plt.bar(x,y)
    ''' 不让标签太密 '''
    if len(x)//6>=20:
        plt.xticks([x[i] for i in range(0,len(x),12)])
    else:
        plt.xticks([x[i] for i in range(0,len(x),6)])
    plt.gcf().autofmt_xdate()
    plt.savefig('./asset/Bar.png')
    plt.show()
    

# 以下调试用代码，请忽略
if __name__ == '__main__':
    data = pd.read_excel('./data/data.xlsx')
    y = data[data.columns[1]]
    x = data[data.columns[0]]
    if str(type(x[0])) != '''<class 'str'>''':
        x = x.apply(lambda i: i.strftime('%Y-%m'))
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
    plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
    fig = plt.figure(figsize=(8.0,5.6))
    plt.xlabel(data.columns[0])
    plt.ylabel(data.columns[1])
    plt.bar(x,y)
    plt.xticks([x[i] for i in range(0,len(x),len(x)//9)])
    plt.gcf().autofmt_xdate()
    plt.savefig('./asset/Bar.png')
    plt.show()
