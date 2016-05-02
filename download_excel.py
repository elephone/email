#coding:utf-8
import  read_excel
import os
import xlwt
import xlrd

file_path = u'F:\经销商反馈424'
paths = [os.path.join(file_path, f) for f in os.listdir(file_path)]
contents =[]



w = xlwt.Workbook()
ws = w.add_sheet('sheet1')  # 创建一个工作表
i = 0
j = 0
ws.write(i, j, u'区域')
ws.write(i, j + 1, u'省份')
ws.write(i, j + 2, u'城市')
ws.write(i, j + 3, u'姓名')
ws.write(i, j + 4, u'性别')
ws.write(i, j + 5, u'电话')
ws.write(i, j + 6, u'活动名称')
ws.write(i, j + 7, u'媒体渠道')
ws.write(i, j + 8, u'数据创建日期')
ws.write(i, j + 9, u'经销商名称')
ws.write(i, j + 10, u'具体型号')
ws.write(i, j + 11, u'呼叫结果')
ws.write(i, j + 12, u'计划购车时间')
ws.write(i, j + 13, u'预约到店时间')
ws.write(i, j + 14, u'回访描述（经销商完成）')
ws.write(i, j + 15, u'客户定级（经销商完成）')
ws.write(i, j + 16, u'是否试驾（经销商完成）')
ws.write(i, j + 17, u'是否下订单（经销商完成）')

for p in paths:
    contents.extend(read_excel.excel_table_byindex(p))

for it in contents:

    i += 1
    j = 0
    for (d, x) in it.items():
        ws.write(i,j,x)
        j += 1
        # print "key:" + d + ",value:" + str(x)

w.save( u'F:/反馈数据424.xls')


def excel_table_byindex(file= u'分发销售线索名单422.xlsx',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list
def open_excel(file= u'分发销售线索名单422.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)