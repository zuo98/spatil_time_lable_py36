# -*- coding: UTF-8 -*-
import func
import os
import datetime
import pandas as pd


# excel数据表路径
ExcelPath = 'D:\\Teacher Song\\spatil_time_label_test\\补充数据测试\\excel'
# 地名库excel数据表
countyExcel = 'D:\\Teacher Song\\spatil_time_label_test\\补充数据测试\\地名库指标库\\地名库.xlsx'
# 输出路径
outPath = 'D:\\Teacher Song\\spatil_time_label_test\\outExcel'

yearList = ['2011', '2012', '2013', '2014', '2015']# 有效年份
attributesList = ['POPYE', 'FIX', 'RESID', 'VPOP', 'CPOP', 'GDP', 'GDP1', 'GDP2', 'GDP3']# 有效字段名
countyList = func.getCountyList(countyExcel)# 获取有效县区名称

startTime = datetime.datetime.now()

log = outPath+"\\log.txt" # 新建一个txt文件收集出现错误的excel表
files = open(log, "w")

fileList = os.listdir(ExcelPath)
StandarDataList = []

for f in fileList:
    if '.xls' in f[-4:]:
        fullPath = os.path.join(ExcelPath, f)
        print(fullPath)
        try:
            startPoint = func.getStartPoint(fullPath) # 获取keyword的行列及其值
            ExcelType = func.getExcelType(startPoint) # 获取excel表的类型
            df = func.clearData(fullPath, startPoint) # 去掉多余的行列
            dataList = func.getStandardData(df, ExcelType, startPoint) #将数据逐条存入dataList
            # 将数据逐条规范化为[year, county, attributes, values],符合要求就存入StandarDataList中
            for data in dataList:
                print(data)
                data = func.clearStandarData(data, countyList, yearList, attributesList)
                print(data)
                if data != []:
                    StandarDataList.append(data)
        except:
            # 上述出错就将出错的excel名称保存在‘log.txt’中
            files.write(fullPath+'\n')

# 将StandarDataList转化为DataFrame形式，并导出一份名为AStandarData的excel表
df = pd.DataFrame(StandarDataList, columns=['year', 'county', 'attributes', 'value'])
df.to_excel(outPath+'\\AStandarData.xls')

# 根据年份，把数据分开，每一年的数据DataFrame通过DataFrame.pivot()方法转化为以county为列，attributes为行，values为值得透视表
# 并每一年导出一个excel，
for y in yearList:
    dfyear = df[df.year == y]
    dfyear = dfyear.pivot(index='county', columns='attributes', values='value')
    print(dfyear)
    dfyear.to_excel(outPath+'\\excelBy{}.xls'.format(y))

files.write(str(datetime.datetime.now()))
files.close()# 在log.txt中写入日期，关闭log.txt文件。

endTime = datetime.datetime.now()
print('use time: {} s'.format(str((endTime-startTime).seconds)))
