# -*- coding:utf-8 -*-
__author__ = 'IanChen'

import xlrd

data = xlrd.open_workbook("工作簿1.xls")
print(data)
table = data.sheets()[0]#通过索引获取

table2 = data.sheet_by_name(u'Sheet1')#通过名称获取
hangshuju=table.row_values(2)
danyuange=hangshuju[3]
lieshuju = table2.col_values(1)
nrows=table.nrows#行数
lieshu=table.ncols#列数
danyuange1=table.cell(0,3).value
cell= table.row(0)[3].value#使用行列索引
print(cell)

#写入

import xlwt
workbook = xlwt.Workbook(encoding = 'utf8')
worksheet = workbook.add_sheet('My Worksheet')#创表
worksheet.write(0, 0, label = 'Row 0, Column 0 Value')
workbook.save('Excel_Workbook.xls')

#修改excel
from xlutils.copy import copy
newWB=copy(data)
newWS=newWB.get_sheet(0)
newWS.write(0,0,"xiugai")
newWB.save("工作簿1.xls")