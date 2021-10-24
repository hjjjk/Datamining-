import os
import csv
import xlwt
import numpy as np
import pandas as pd

def access_data(path, path_1):
    df = pd.read_excel(path).iloc[:,1:7]
    table_info = np.array(df-1000)
    table_info[np.isnan(table_info)]=0
    table_info_list = table_info.tolist()
    list_num = len(table_info_list)

    for i in range(list_num):
        while 0 in table_info_list[i]:
            table_info_list[i].remove(0)
    for a in range(list_num):
        if len(table_info_list[a])== 2:
            table_info_list[a] = str(table_info_list[a][0])+';'+str(table_info_list[a][1])
        elif len(table_info_list[a])==3:
            table_info_list[a] = str(table_info_list[a][0]) + ';' + str(table_info_list[a][1]) + ';' +str(table_info_list[a][2])
        elif len(table_info_list[a])==4:
            table_info_list[a] = str(table_info_list[a][0]) + ';' + str(table_info_list[a][1]) + ';' +str(table_info_list[a][2])+ ';' +str(table_info_list[a][3])
        elif len(table_info_list[a])==5:
            table_info_list[a] = str(table_info_list[a][0]) + ';' + str(table_info_list[a][1]) + ';' +str(table_info_list[a][2])+ ';' +str(table_info_list[a][3])+ ';' +str(table_info_list[a][4])
        elif len(table_info_list[a])==6:
            table_info_list[a] = str(table_info_list[a][0]) + ';' + str(table_info_list[a][1]) + ';' +str(table_info_list[a][2])+ ';' +str(table_info_list[a][3])+ ';' +str(table_info_list[a][4])+ ';' +str(table_info_list[a][5])

    table_info_1 = np.array(df)
    table_info_1[np.isnan(table_info_1)]=0
    table_info_list_1 = table_info_1.tolist()
    for b in range(list_num):
        while 0 in table_info_list_1[b]:
            table_info_list_1[b].remove(0)

    for c in range(list_num):
        if len(table_info_list_1[c])== 2:
            table_info_list_1[c] = str(table_info_list_1[c][0])+';'+str(table_info_list_1[c][1])
        elif len(table_info_list_1[c])==3:
            table_info_list_1[c] = str(table_info_list_1[c][0]) + ';' + str(table_info_list_1[c][1]) + ';' +str(table_info_list_1[c][2])
        elif len(table_info_list_1[c])==4:
            table_info_list_1[c] = str(table_info_list_1[c][0]) + ';' + str(table_info_list_1[c][1]) + ';' +str(table_info_list_1[c][2])+ ';' +str(table_info_list_1[c][3])
        elif len(table_info_list_1[c])==5:
            table_info_list_1[c] = str(table_info_list_1[c][0]) + ';' + str(table_info_list_1[c][1]) + ';' +str(table_info_list_1[c][2])+ ';' +str(table_info_list_1[c][3])+ ';' +str(table_info_list_1[c][4])
        elif len(table_info_list_1[c])==6:
            table_info_list_1[c] = str(table_info_list_1[c][0]) + ';' + str(table_info_list_1[c][1]) + ';' +str(table_info_list_1[c][2])+ ';' +str(table_info_list_1[c][3])+ ';' +str(table_info_list_1[c][4])+ ';' +str(table_info_list_1[c][5])


    table_info_2 = np.array(df)
    table_info_2[np.isnan(table_info_2)]=0
    table_info_list_2 = table_info_2.tolist()
    for d in range(list_num):
        while 0 in table_info_list_2[d]:
            table_info_list_2[d].remove(0)

    table_lable_list = []
    for e in table_info_list_2:
        for f in e:
            table_lable_list.append(f)

    table_lable_list_max = np.array(table_lable_list)
    table_lable_list_max = np.unique(table_lable_list_max)
    table_lable_list_max_sort = [g+1 for g in range(len(table_lable_list_max))]



    table_head = ["2009-2014","组织","重新编码","替换后"]
    workbook = xlwt.Workbook(encoding = 'utf-8')
    xlsheet = workbook.add_sheet("五年数据",cell_overwrite_ok=True)
    # 写表头
    headlen = len(table_head)
    for h in range(headlen):
        xlsheet.write(0, h, table_head[h])

    for j in range(list_num):
        xlsheet.write(j+1,0,str(table_info_list_1[j]))
        xlsheet.write(j+1,3,str(table_info_list[j]))

    for k in range(len(table_lable_list_max)):
        xlsheet.write(k+1,1,int(table_lable_list_max[k]))
        xlsheet.write(k+1,2,int(table_lable_list_max_sort[k]))

    workbook.save(path_1)




path_pack = r'D:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/'
pack_age = os.listdir(path_pack)[1:]
for i in pack_age:
    path = 'D:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/{0}/名称转编号.xls'.format(i)
    path_1 = 'D:/挖掘/实验一：数据及数据预处理/电子信息产业原始数据/{0}/2009-2014合作编码(两两组合).xls'.format(i)

    access_data(path, path_1)

