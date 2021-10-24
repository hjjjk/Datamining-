"""
选择拥有专利的企业
"""
import xlwt
import numpy as np
import pandas as pd

path = 'F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/六年数据.xls'
data_info = pd.read_excel(path)['申请（专利权）人']
data_1 = np.array(data_info).tolist()
data_2 = [i.split(";") for i in data_1]

data_num = len(data_2)
data_index = [len(data_2[i]) for i in range(data_num)]
data_3 = [data_2[i] for i in range(data_num)]
# 将所有数据放入一个列表中
a = []
for i in range(data_num):
    a = a+data_3[i]
c = np.array(a)
d = np.unique(c).tolist()



workbook = xlwt.Workbook(encoding = 'utf-8')
xlsheet = workbook.add_sheet("专利归属方",cell_overwrite_ok=True)
xlsheet.write(0, 0, '申请（专利权）人')
workbook.save('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/6电子计算机外部设备/专利人_1.xls')
for i in d:
    print(d)


workbook = xlwt.Workbook(encoding = 'utf-8')
xlsheet = workbook.add_sheet("专利归属方",cell_overwrite_ok=True)
xlsheet.write(0, 0, '申请（专利权）人')

for rows in range(len(d)):

    xlsheet.write(rows + 1, d[rows])

workbook.save('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/6电子计算机外部设备/专利人.xls')
