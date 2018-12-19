# -*- coding: utf-8 -*-
from dbfread import DBF
from xlutils.copy import copy
import os
import xlrd


shpPath = r'D:\Teacher Song\spatil_time_label_test\shp'
countyExcel = r'D:\Teacher Song\spatil_time_label_test\补充数据测试\地名库指标库\地名库补充后.xls'

rb = xlrd.open_workbook(countyExcel)
table = rb.sheet_by_index(0)
rows = table.nrows
print(rows)
cols = table.ncols
nameList = []
for i in range(1, rows):
    for j in range(cols):
        name = table.cell(i, j).value
        name = name.replace(' ', '')
        if name != '':
            nameList.append(name)

dbfFile = os.listdir(shpPath)
addName = []
for dbf in dbfFile:
    if '.dbf' in dbf[-4:]:
        dbfFilePath = os.path.join(shpPath, dbf)
        print(dbfFilePath)
        dbftable = DBF(dbfFilePath, char_decode_errors='ignore', raw=True)
        for record in dbftable:
            name = record['NAME'].decode('utf-8')
            name = name.replace(' ', '')
            if name != '' and (name not in nameList) and (name not in addName):
                print(name)
                addName.append(name)
# print(addName)
wb = copy(rb)
ws = wb.get_sheet(0)
cont = 0
for name in addName:
    ws.write(rows+cont, 0, name)
    cont += 1
wb.save(r'D:\Teacher Song\spatil_time_label_test\补充数据测试\地名库指标库\地名库补充后.xls')
