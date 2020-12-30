# -*- coding: utf-8 -*-
# @Time    : 2020/9/28 22:26
# @Author  : Sheldon.Tian
# @Email   : ty24089@163.com
# @File    : TYReadExc.py
# @Software:

import numpy as np
import pandas as pd


# 读取文件
def read_from_exc(path):
    df = pd.read_excel(r'~/Documents/12月份绩效-田宇.xlsx', sheet_name='Sheet1', parse_dates=True)
    # print(df.columns)
    # print(df.loc[[0], ['Unnamed: 2']].values)
    # print(df.loc[[21], ['Unnamed: 1']].values)
    print(df.iloc[21])
    data = df.values
    # print(data)
    return data


# 解析数组
def analusis(data:pd.DataFrame):
    # 读取二维数组的值
    for val in data:
        if '被考核人' in val:
            # 读取一维数组的值
            for item in val:
                print("获取到的所有的值：\n{0}".format(item))


data = read_from_exc('')
# analusis(data)

# 绩效得分总计