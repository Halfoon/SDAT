# SDAT- A Simple Data Analysis Tool

## 介绍
詹骋昊 20377085 的大学计算机基础（ 2021 年春）-大作业第一题

## 运行
`python gui.py`
or
`python main.py`

## 已实现功能
1. 用户选择数据，多格式数据文件支持（csv, excel(xls and xlsx), txt)
2. 单变量数据支持，基本描述性统计量分析并展示
3. 数据可视化（柱状图，折线图）并保存文件于指定位置
4. gui界面展示导入数据表格
5. gui界面信息提示
6. 报告（Markdown格式）生成

## 依赖
- tk
- pandas
- matlibplot
- openpyxl
- xlrd
- tksheet
- tabulate
以上库 PyPl 均有

## 注意事项
必须保证源文件路径有名为"asset"的文件夹，否则无法保存图片

## 已知问题
1. 在数据可视化（使用plt.show()）时 gui 界面缩放比例会改变，部分内容显示出错，怀疑与 Windows10 的缩放机制有关