"""
根据对应的公司名索引来替换文字
"""
import os
import xlwt
import numpy as np
import tensorflow as tf
import pandas as pd


def make_data(path):
    df = pd.read_excel(path+'专利人.xls')
    df_array = np.array(df).tolist()
    df_index = dict(df_array)

    df_1 = pd.read_excel(path+'六年数据.xls')['申请（专利权）人']
    df_1_array = np.array(df_1).tolist()
    # 每一项专利的申请人名单
    data_1 = [i.split(";") for i in df_1_array]
    data_1_index = [i for i in range(len(data_1))]

    workbook = xlwt.Workbook(encoding = 'utf-8')
    xlsheet = workbook.add_sheet("excel写入练习",cell_overwrite_ok=True)
    xlsheet.write(0, 0, '专利方名称')
    xlsheet.write(0, 1, '专利方转编号')
    for i in data_1_index:
        xlsheet.write(i+1,0,str(data_1[i]))
        for b in range(len(data_1[i])):
            up_name = data_1[i][b]
            up_name_num =df_index[up_name]
            xlsheet.write(i+1,b+1,up_name_num)

    workbook.save(path+'名称转编号_x.xls')

path = 'D:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/7电子计算机数据处理及应用/'
make_data(path)
