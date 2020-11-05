# -*- coding: utf-8 -*-
"""
Created on Tue Oct  6 10:28:27 2020

@author: yy
@description: gspread使用
"""

"""读取【版权数据】"""
import gspread

gc = gspread.oauth()   # 获取谷歌文档访问凭据
ssh = gc.open("版权数据")   # 通过文档名打开
ssh2 = gc.open_by_key('1-aUVDHtdDCncPz-548lPplBjAkptH7X4pRx7_72WbBA')   # 通过文档key打开
ssh3 = gc.open_by_url('https://docs.google.com/spreadsheets/d/1-aUVDHtdDCncPz-548lPplBjAkptH7X4pRx7_72WbBA/edit#gid=0')


"""选择某一个sheet"""
worksheet_list = ssh.worksheets()   # 获取文件中的所有sheets信息

for ws in worksheet_list:   # 查看各sheets的属性信息
    """ 打印结果：title就是sheet名；index为sheet的索引，其实就是打开worksheet后下方的次序
        图表 0
        进审 1
        总体 2
        一组 3
        二组 4
        三组 5
        四组 6
        新人数据 7
        工作表9 8
    """
    print(ws._properties['title'], ws._properties['index'])

zongti = ssh.worksheet("总体")   # 通过sheet名称获取
zongti2 = ssh.get_worksheet(2)   # 通过sheet的index获取


"""读取单元格内容"""
val = zongti.acell('B2').value   # 通过单元格名获取值
val2 = zongti.cell(2,2).value   # 通过行列号获取值，注意这里行列号是从1开始！

cell = zongti.acell('B2', value_render_option='FORMULA').value   # 若单元格有计算公式，获取计算公式
cell1 = zongti.cell(2, 2, value_render_option='FORMULA').value   # 若单元格有计算公式，获取计算公式


"""读取行/列内容"""
row_val = zongti.row_values(1)   # 获取首行值
col_val = zongti.col_values(1)   # 获取首列值


"""读取某范围的单元格值"""
vals = zongti.get_all_values()   # 获取所有值，返回为列表的列表，每一行内容为一个列表元素，如[[行1],[行2]...]
vals2 = zongti.get_all_records()   # 获取所有值，返回为字典的列表，每一行内容为一个字典，字典的key为首行字段名

vals3 = zongti.get('A1:B5')   # 获取连续范围的单元格值
vals4 = zongti.batch_get(['A1:B5', 'A544'])   # 获取不连续范围的单元格值


"""写入/修改单元格值"""
zongti.update('A545', '2020-10-6')   # 通过单元格名称，写入/修改单元格内值
zongti.update_cell(545, 1, '2020-10-6')   # 通过行列号，写入/修改单元格内值
zongti.update('B545:D545', [['1', '2', '43']])   # 连续范围写入/修改
zongti.batch_update([{   # 不连续范围写入/修改
                      'range': 'B545:D545', 
                      'values': [['1', '2', '43']]},
                     {
                      'range': 'B545:B547', 
                      'values': [['1'], ['2'], ['43']]}
                     ])
                    

"""查找定位"""
cell = zongti.find("人效")   # 找到的结果为首个匹配成功的单元格
print("Found something at %s行%s列" % (cell.row, cell.col))

cells = zongti.findall("人效")   # 找到所有匹配成功的单元格，返回为单元格的列表
for c in cells:
    print("Found something at %s行%s列" % (c.row, c.col))

import re
pattern = re.compile(r'审出')
cell = zongti.find(pattern)


"""使用pandas读入表格"""
import pandas as pd
df = pd.DataFrame(zongti.get_all_records())   # 读入dataframe

#"""将整个dataframe写入google spreasheet"""
#import pandas as pd
#zongti.update([df.columns.values.tolist()] + df.values.tolist())   # 字段名+值

