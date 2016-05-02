#coding:utf-8
import  read_excel
import os
import xlwt

file_path = u'F:\经销商反馈424'
paths = [os.path.join(file_path, f) for f in os.listdir(file_path)]
contents =[]

for p in paths:
    contents.extend(read_excel.excel_table_byindex(p))

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


for it in contents:

    i += 1
    j = 0
    # for (d, x) in it.items():
    #     ws.write(i,j,str(x))
    #     j += 1
    #     # print "key:" + d + ",value:" + str(x)
    try:
        ws.write(i, j, it[u'区域'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'省份'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'城市'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'姓名'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'性别'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'电话'])
    except:
        try:
            ws.write(i,j,it[u'手机'])
        except:
            try:
                ws.write(i,j,it[u'手机号码'])
            except Exception,ex:
                print Exception,':',ex.message
    try:
        j += 1
        ws.write(i, j , it[u'活动名称'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j , it[u'媒体渠道'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'数据创建日期'])
    except:
        try:
            ws.write(i, j, it[u'创建时间'])
        except Exception,ex:
            print Exception,':',ex.message
    try:
        j += 1
        ws.write(i, j, it[u'经销商名称'])

    except:
        try:
            ws.write(i,j,it[u'经销商'])
        except Exception, ex:
            print Exception, ':', ex.message
    try:
        j += 1
        ws.write(i, j , it[u'具体型号'])

    except:
        try:
            ws.write(i,j,it[u'意向车型'])
        except Exception, ex:
            print Exception, ':', ex.message
    try:
        j += 1
        ws.write(i, j , it[u'呼叫结果'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'计划购车时间'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j , it[u'预约到店时间'])
    except:
        print ''
    try:
        j += 1
        ws.write(i, j, it[u'回访描述（经销商完成）'])
    except:
        try:
            ws.write(i, j, it[u'回访描述'])
        except Exception,ex:
            print Exception,':',ex.message
    try:
        j += 1
        ws.write(i, j , it[u'客户定级（经销商完成）'])
    except:
        try:
            ws.write(i,j,it[u'客户定级'])
        except Exception,ex:
            print Exception,':',ex.message
    try:
        j += 1
        ws.write(i, j, it[u'是否试驾（经销商完成）'])
    except:
        try:
            ws.write(i, j, it[u'是否试驾'])
        except Exception, ex:
            print Exception, ':', ex.message
    try:
        j += 1
        ws.write(i, j, it[u'是否下订单（经销商完成）'])
    except:
        try:
            ws.write(i, j, it[u'是否下订单'])
        except Exception, ex:
            print Exception, ':', ex.message
w.save( u'F:/反馈数据424.xls')


