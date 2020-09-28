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
    df = pd.read_excel(r'E:\11月绩效-田宇.xlsx', sheet_name='绩效样表', parse_dates=True)
    data = df.values
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
analusis(data)