# -*- coding: utf-8 -*-
from xlutils.copy import copy
import os
import xlrd

# 原始的未标注的excel表路径
excelPath = r'D:\Teacher Song\spatil_time_label_test\补充数据测试\excel'
sameNameExcel = r'D:\Teacher Song\spatil_time_label_test\补充数据测试\地名库指标库\存在同名的县区.xlsx'

rb = xlrd.open_workbook(sameNameExcel)
table = rb.sheet_by_index(0)
rows = table.nrows
cols = table.ncols

sameNameList = []
for i in range(1, rows):
    for j in range(cols):
        name = table.cell(i, j).value
        name = name.replace(' ', '')
        name = name.replace('\n', '')
        if name != '':
            print(name)
            sameNameList.append(name)

excelList = os.listdir(excelPath)
for excel in excelList:
    if '.xls' in excel[-4:]:
        path = os.path.join(excelPath, excel)
        print(path)
        wb = xlrd.open_workbook(path)
        table = wb.sheet_by_index(0)
        rows = table.nrows
        cols = table.ncols
        cwb = copy(wb)
        ctable = cwb.get_sheet(0)
        for i in range(rows):
            for j in range(cols):
                value = table.cell(i, j).value
                if isinstance(value, str):
                    value = value.replace(' ', '')
                    value = value.replace('\n', '')
                    if value in sameNameList:
                        print(value+'需要加地市名')
                        ctable.write(i, j, 'addcity'+value)
                        cwb.save(path)
