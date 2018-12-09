# -*- coding: UTF-8 -*-
import pandas as pd
import xlrd
import re


# 获取地名库Excel表中的所有地名，返回包含所有县区名的list
def getCountyList(nameExcel):
    countyList = []
    df = pd.read_excel(nameExcel)
    seriesNames = df.columns.values.tolist()
    rows = df.shape[0]
    for name in seriesNames:
        series = df[name]
        for j in range(rows):
            county = series[j]
            if isinstance(county, str):
                countyList.append(county)
    return countyList


# 获取标注的行列号 即标注饭后list[rows, cols, keyWord]
def getStartPoint(dataExcel, searchWith=6):
    rows = 0
    cols = 0
    keyWord = ''
    xl = xlrd.open_workbook(dataExcel)
    table = xl.sheet_by_name('sheet1')
    cell = 0
    while cell < searchWith:
        for i in range(cell+1):
            value = table.cell(i, cell).value
            if isinstance(value, str):
                if ('mc' in value) or ('at' in value) or ('yr' in value):
                    rows = i
                    cols = cell
                    keyWord = value
                    break
            value = table.cell(cell, i).value
            if isinstance(value, str):
                if ('mc' in value) or ('at' in value) or ('yr' in value):
                    rows = cell
                    cols = i
                    keyWord = value
                    break
        if keyWord != '':
            break
        cell = cell + 1
    if keyWord == '':
        int('error')
    return [rows, cols, keyWord]


# 通过keyword，判断属于那种情况，返回情况类型，若都不属于，则让他强行报错，在主程序中设置try..except..收集错误excel
def getExcelType(startPoint):
    ExcelType = ''
    if str(re.match("[m][c][1-9][0-9]{3}[a][t]", startPoint[2])) != 'None':
        ExcelType = 'typeOne'
    elif str(re.match("[m][c][\u4e00-\u9fa5]{2,}[y][r][a][t]", startPoint[2])) != 'None':
        ExcelType = 'typeTwo'
    elif str(re.match("[m][c][y][r]", startPoint[2])) != 'None':
        ExcelType = 'typeThree'
    else:
        int('error')
    return ExcelType


# 删去未标注的列或行，返回DataFrame
def clearData(dataExcel, startPoint):
    df = pd.read_excel(dataExcel, skiprows=startPoint[0])
    columns = df.columns
    for column in columns:
        if isinstance(column, str):
            if 'Unname' in column:
                df.drop(columns=[column], inplace=True)
    df.dropna(axis='index', how='any', inplace=True)
    df.set_index(startPoint[2], inplace=True)
    return df


# 传入已删除多余行列的DataFrame，将里面的数据逐条去取出，储存data=[]中，并返回。
# 这里的逐条数据并不是[year, county, attributes, values]顺序排列，只能确定values位置正确。
def getStandardData(DataFrame, ExcelType, startPoint):
    data = []
    if ExcelType == 'typeOne':
        key = startPoint[2][2:-2]
        for index, row in DataFrame.iterrows():
            for col_name in DataFrame.columns:
                data.append([key, index, col_name, row[col_name]])
    if ExcelType == 'typeTwo':
        key = startPoint[2][2:-4]
        for index, row in DataFrame.iterrows():
            for col_name in DataFrame.columns:
                data.append([key, index, col_name, row[col_name]])
    if ExcelType == 'typeThree':
        key = startPoint[2][4:]
        for index, row in DataFrame.iterrows():
            for col_name in DataFrame.columns:
                data.append([key, index, col_name, row[col_name]])
    return data


# 将上面的数据逐条转化为严格的[year, county, attributes, values]顺序，并筛掉不符合的数据，最后返回
def clearStandarData(standarData, countyList, yearList, attributesList):
    clearStandarData = ['year', 'county', 'attributes', 'value']
    for data in standarData[:3]:
        if isinstance(data, str):
            data = data.replace('\n', '')
            data = data.replace(' ', '')
        else:
            data = str(data)
            data = data.replace('\n', '')
        if data in yearList:
            clearStandarData[0] = data
        elif data in countyList:
            clearStandarData[1] = data
        elif data in attributesList:
            clearStandarData[2] = data

    if isinstance(standarData[3], str):
        clearStandarData[3] = float(standarData[3].replace('\n', ''))
    else:
        clearStandarData[3] = standarData[3]
    if clearStandarData[0] != 'year' and clearStandarData[1] != 'county' and clearStandarData[2] != 'attributes' and clearStandarData[3] != 'value':
        return clearStandarData
    else:
        return []
