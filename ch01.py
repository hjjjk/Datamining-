"""
合并数据
"""
import os
import csv
import xlwt
import numpy as np
import pandas as pd

def make_data(all_data):
    table_head = ['申请号','名称','主分类号','分类号','申请（专利权）人',	'发明（设计）人','公开（公告）日'	,'公开（公告）号',	'专利代理机构','代理人','申请日','地址','摘要'	,'国省代码']
    workbook = xlwt.Workbook(encoding = 'utf-8')
    xlsheet = workbook.add_sheet("excel写入练习",cell_overwrite_ok=True)
    # 写表头
    headlen = len(table_head)
    for i in range(headlen):
        xlsheet.write(0, i, table_head[i])
    # 获取有多少条数据
    all_data_num = len(all_data)
    for row in range(all_data_num):
        for col in range(14):
            xlsheet.write(row+1,col,all_data[row][col])

    workbook.save('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/六年数据.xls')



table_info_1 = pd.read_excel('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/2009_1.xls')
table_info_2 = pd.read_excel('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/2010_1.xls')
table_info_3 = pd.read_excel('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/2011_1.xls')
table_info_4 = pd.read_excel('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/2012_1.xls')
table_info_5 = pd.read_excel('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/2013_1.xls')
table_info_6 = pd.read_excel('F:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/2014_1.xls')

table_info = pd.concat((table_info_1,table_info_2,table_info_3,table_info_4,table_info_5,table_info_6))
table_info_x = np.array(table_info)
all_data = table_info_x.tolist()
make_data(all_data)








# 写入数据
# workbook = xlsxwriter.Workbook(target_xls)  # 创建了一个名字叫做3.xlsx ， Excel表格文件
# worksheet = workbook.add_worksheet()  # 建立sheet,
# font = workbook.add_format({"font_size": 14})  # 表格中值（字体）的大小
# for i in range(len(data)):  # 从data列表中读取数据
#     for j in range(len(data[i])):
#         worksheet.write(i, j, data[i][j], font)
# # 关闭文件流
# workbook.close()


# table_info = pd.read_excel(source_path)
# table_info_1 = pd.read_excel(source_path_1)