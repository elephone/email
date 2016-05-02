#-*- coding: utf8 -*-
from xlrd import open_workbook
from xlutils.copy import copy
import read_excel

rb = open_workbook(u'分发销售线索名单422.xls')

#通过sheet_by_index()获取的sheet没有write()方法
rs = rb.sheet_by_index(0)


wb = copy(rb)

#通过get_sheet()获取的sheet有write()方法
ws = wb.get_sheet(0)


nrows = rs.nrows #行数
ncols = rs.ncols #列数
# colnames =  table.row_values(1) #某一行数据
# list =[]
# for rownum in range(1,nrows):
#      row = table.row_values(rownum)
#      if row:
#          app = {}
#          for i in range(len(colnames)):
#
#            app[colnames[i]] = row[i]
#          list.append(app)
for i in  range(0,ncols,1):
    ws.write(1,i,'')


wb.save(u'分发销售线索名单422.xls')