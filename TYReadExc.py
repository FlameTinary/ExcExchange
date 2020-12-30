# -*- coding: utf-8 -*-
# @Time    : 2020/9/28 22:26
# @Author  : Sheldon.Tian
# @Email   : ty24089@163.com
# @File    : TYReadExc.py
# @Software:

import pandas as pd
import os
from openpyxl import load_workbook

# 这里需要手动改为绩效表存放位置
dirPath = r'E:\技术部12月绩效(1)\AI项目组12月绩效\AI项目组12月绩效\前端'

# 这里是生成表的名称
newFileName = r'绩效总表'


def getFiles():
    fileNames = os.listdir(dirPath)
    allfile = []
    for file in fileNames:
        allfile.append(dirPath + '\\' + file)
    return allfile


# 从单独文件中获取需要的内容
def read_from_exc(path):
    df = pd.read_excel(path, parse_dates=True)
    name = ''
    score = ''
    # 获取每一行的index和对应的内容
    for index, row in df.iterrows():
        # 去除NaN
        row = row.dropna(axis=0)
        # 获取姓名
        if '被考核人' in row.values:
            name = row.values[2]
            # print(row.values[2])

        if '绩效得分总计' in row.values:
            score = row.values[1]
            # print(row.values[1])
    return [name, score]


# print(read_from_exc(''))

# 将单独文件中的数据汇总为DataFrame
def makeFrame():

    names = []
    scores = []
    for fileName in getFiles():
        # 获取需要的值
        res = read_from_exc(fileName)
        names.append(res[0])
        scores.append(res[1])
    nameSeries = pd.Series(names)
    scoresSeries = pd.Series(scores)
    d = {'姓名': nameSeries, '绩效': scoresSeries}
    df = pd.DataFrame(d)
    return df


# 写入文件
def dataFrameToFile(df: pd.DataFrame):
    if not os.path.exists(r'E:\绩效总表.xlsx'):
        nan_excle = pd.DataFrame()
        nan_excle.to_excel(r'E:\绩效总表.xlsx')

    df1 = pd.DataFrame(pd.read_excel(r'E:\绩效总表.xlsx'))  # 读取原数据文件和表
    writer = pd.ExcelWriter(r'E:\绩效总表.xlsx', engine='openpyxl')
    book = load_workbook(r'E:\绩效总表.xlsx')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df_rows = df1.shape[0]  # 获取原数据的行数
    df.to_excel(writer, startrow=df_rows + 1, index=False, header=False)  # 将数据写入excel中的aa表,从第一个空行开始写
    writer.save()  # 保存


dataFrameToFile(makeFrame())

# 新需求
# 总表去重功能
#