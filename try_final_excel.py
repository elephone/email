#-*- coding: utf8 -*-
import sys
import read_excel
from pyExcelerator import *
from xlutils.copy import copy
import win32com.client as win32
import os
import xlwt
import datetime,time
import smtplib
import email.MIMEBase  # import MIMEBase
import os.path
import time



def get_all_customs_data(path):
    #get data
    title = u'经销商名称'
    return read_excel.excel_table_byindex(path)

def get_the_main_address(custom_shop_name):
    title = u'经销商名称'
    email_address = u'邮箱'
    contents = read_excel.excel_table_byindex(u'撼路者经销商通讯录.xlsx',0,0)
    for it in contents:
        if(it[title] == custom_shop_name):
            return it[email_address]
    # try:
    #     print ''
    #     #no find the shop name in the whole
    # except:
    #     print u''
    print u'!!!!error'+ custom_shop_name
    return u'!!!!error'+ custom_shop_name

def get_the_copy_addresses(custom_area):
    if (custom_area == u'东区'):
        return '''
gwang4@jmc.com.cn;
wwang16@jmc.com.cn;
jhe3@jmc.com.cn;
mhu4@jmc.com.cn;
jxiao9@jmc.com.cn;
fwei1@jmc.com.cn;
chong1@jmc.com.cn;
xdeng9@jmc.com.cn;
Stephanie.Ji@wunderman.com;
amy.sun@wunderman.com'''
    elif(custom_area == u'南区'):
        return '''
jyu13@jmc.com.cn;
zwu8@jmc.com.cn;
lcui@jmc.com.cn;
qyu2@jmc.com.cn;
bliu16@jmc.com.cn;
jwu25@jmc.com.cn;
wzhao8@jmc.com.cn;
fwei1@jmc.com.cn;
chong1@jmc.com.cn;
xdeng9@jmc.com.cn;
Stephanie.Ji@wunderman.com;
amy.sun@wunderman.com'''

    elif(custom_area == u'西区'):
        return '''
hwang13@jmc.com.cn;
hchen11@jmc.com.cn;
xzhang45@jmc.com.cn;
wlv3@jmc.com.cn;
jli55@jmc.com.cn;
fwei1@jmc.com.cn;
chong1@jmc.com.cn;
xdeng9@jmc.com.cn;
Stephanie.Ji@wunderman.com;
amy.sun@wunderman.com'''

    elif(custom_area == u'北区'):
        return '''
lcai3@jmc.com.cn;
yli41@jmc.com.cn;
mpeng3@jmc.com.cn;
dzhou3@jmc.com.cn;
dwei3@jmc.com.cn;
fwei1@jmc.com.cn;
chong1@jmc.com.cn;
xdeng9@jmc.com.cn;
Stephanie.Ji@wunderman.com;
amy.sun@wunderman.com'''

    elif(custom_area == u'中区'):
        return '''
gzhang2@jmc.com.cn;
tyang1@jmc.com.cn;
tzeng1@jmc.com.cn;
nwang3@jmc.com.cn;
zwan16@jmc.com.cn;
fwei1@jmc.com.cn;
chong1@jmc.com.cn;
xdeng9@jmc.com.cn;
Stephanie.Ji@wunderman.com;
amy.sun@wunderman.com'''

def create_package(tmp_shop_name=0):
    path = u'f:/a/' + tmp_shop_name

    try:
        os.mkdir(path)
    except:
        print u'创建文件夹失败   --->'+ tmp_shop_name

def create_excel_file(tmp_shop_name=0,excel_contents=0):
    path = u'f:/a/'

    w = xlwt.Workbook()
    ws = w.add_sheet(tmp_shop_name)  # 创建一个工作表
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
    for it in excel_contents:
        i += 1
        j = 0
        ws.write(i, j, it[u'区域'])
        ws.write(i, j + 1, it[u'省份'])
        ws.write(i, j + 2, it[u'城市'])
        ws.write(i, j + 3, it[u'姓名'])
        ws.write(i, j + 4, it[u'性别'])
        ws.write(i, j + 5, it[u'电话'])
        ws.write(i, j + 6, it[u'活动名称'])
        ws.write(i, j + 7, it[u'媒体渠道'])
        ws.write(i, j + 8, it[u'数据创建日期'])
        ws.write(i, j + 9, it[u'经销商名称'])
        ws.write(i, j + 10, it[u'具体型号'])
        ws.write(i, j + 11, it[u'呼叫结果'])
        ws.write(i, j + 12, it[u'计划购车时间'])
        ws.write(i, j + 13, it[u'预约到店时间'])
    now = datetime.date.today().strftime("%m%d")
    w.save(path + '/'+tmp_shop_name+u'分发销售线索名单'+now+u'.xls')
    return path+ '/'+tmp_shop_name+u'分发销售线索名单'+now+'.xls'

def outlook(excel_inf_num,to_address='',cc_address='',tmp_shop_name=''):
    now = datetime.date.today()
    end_send_time = now + datetime.timedelta(days=5)
    # app= 'Outlook'
    # olook = win32.gencache.EnsureDispatch("%s.Application" % app)
    # mail=olook.CreateItem(win32.constants.olMailItem)
    olMailItem = 0x0
    obj = win32.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = u'{0}_撼路者销售线索跟进   {1}'.format(now,tmp_shop_name)
    newMail.Body =u'''各位撼路者经销商同仁 您好，

我们是Wunderman的CRM团队，目前正在以电访的形式对每日新增撼路者销售线索进行甄别，今日获取到的有效销售线索如附件。鉴于线索时效性，请各位务必在以下规定时间内完成跟进，并反馈追踪结果。
{0} 分发{1} 条销售线索，烦请各位及时完成跟进，务必将有效线索录入DMS系统，并将追踪结果于{2} 上午10点之前反馈到邮箱：data.intern@yrbrands.com
如有任何问题，烦请及时与我们联系，谢谢！

江铃福特
 '''.format(now,excel_inf_num,end_send_time)
    newMail.To = to_address
    newMail.CC = cc_address
    newMail.Save()


def stmp(server,excel_inf_num,to_address,cc_address,tmp_shop_name='',file_path=u''):
    now = datetime.date.today()
    end_send_time = datetime.date.today() - datetime.timedelta(days=datetime.date.today().weekday()) + datetime.timedelta(days=7)
    file_name = file_path  # 附件名

    # 构造MIMEMultipart对象做为根容器
    main_msg = email.MIMEMultipart.MIMEMultipart()

    # 构造MIMEText对象做为邮件显示内容并附加到根容器
    text_msg = email.MIMEText.MIMEText(u'''各位撼路者经销商同仁 您好，

我们是Wunderman的CRM团队，目前正在以电访的形式对每日新增撼路者销售线索进行甄别，今日获取到的有效销售线索如附件。鉴于线索时效性，请各位务必在以下规定时间内完成跟进，并反馈追踪结果。
{0} 分发{1} 条销售线索，烦请各位及时完成跟进，务必将有效线索录入DMS系统，并将追踪结果于{2}上午 10点之前反馈到邮箱：data.intern@yrbrands.com
如有任何问题，烦请及时与我们联系，谢谢！

江铃福特

 '''.format(now,excel_inf_num,end_send_time), _charset="utf-8")

    main_msg.attach(text_msg)

    # 构造MIMEBase对象做为文件附件内容并附加到根容器
    contype = 'application/octet-stream'
    maintype, subtype = contype.split('/', 1)

    ## 读入文件内容并格式化 [方式2]－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    data = open(file_name, 'rb')
    file_msg = email.MIMEBase.MIMEBase(maintype, subtype)
    file_msg.set_payload(data.read())
    data.close()
    email.Encoders.encode_base64(file_msg)  # 把附件编码
    # －－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    ## 设置附件头
    basename = os.path.basename(file_name).encode('utf-8')
    file_msg.add_header('Content-Disposition', 'attachment', filename=basename)  # 修改邮件头
    main_msg.attach(file_msg)

    # 设置根容器属性
    main_msg['From'] = "Datainsights.Intern@wunderman.com"
    main_msg['To'] = to_address#TODO
    main_msg['CC'] = cc_address#cc_address#TODO
    main_msg['Subject'] = u'{0}_撼路者销售线索跟进   致{1}'.format(now,tmp_shop_name)
    main_msg['Date'] = email.Utils.formatdate()

    # 得到格式化后的完整文本
    fullText = main_msg.as_string()
    To = []
    From = "data.intern@yrbrands.com"
    # To = cc_address.split(';')
    # To.append(to_address)
    To.append("2420088103@qq.com")
    To.append("574307361@qq.com")
    To.append("1162835933@qq.com")
    # To = ['574307361@qq.com','2420088103@qq.com','1162835933@qq.com','754490227@qq.com','3054637@qq.com','2127861258@qq.com','1839978494@qq.com','574307361@qq.com','2420088103@qq.com']
    # 用smtp发送邮件
    for  only_one_address in To:
        try:
            server.timeout = 5000
            print only_one_address
            server.sendmail(From, only_one_address, fullText)
            time.sleep(2)

        except smtplib.SMTPException, ex:
            print smtplib.SMTPException, ex
            print u'发件失败'

            # 邮箱服务器的连接,用户的验证
            server = smtplib.SMTP("smtp.office365.com", 587)
            # server.set_debuglevel(1)#显示编译信息
            # 向mail发送SMTP "ehlo" 命令
            server.ehlo()
            # 启动TLS模式，mail要求
            server.starttls()
            server.login('data.intern@yrbrands.com', 'Password2015')  # 仅smtp服务器需要验证时
            server.sendmail(From, only_one_address, fullText)
        finally:
            # server.quit()
            # server.login('data.intern@yrbrands.com', 'Password2015')  # 仅smtp服务器需要验证时
            print tmp_shop_name + u'发送成功'





if __name__=="__main__":
    now = datetime.date.today().strftime("%m%d")
    path = u'分发销售线索名单'+now+'.xlsx'#customs' information excel file path
    num = 0
    #read the 'xlsx' file,get the whole data
    customs = get_all_customs_data(path)#save by iterator
    #find all the email addresses mapping  customs
    excel_content = []#the excel should to save (for the only one shop name)
    tmp_shop_name =''
    title = u'经销商名称'
    city = u'省份'
    area = u'区域'
    tmp_custom = 0
    # 邮箱服务器的连接,用户的验证
    server = smtplib.SMTP("smtp.office365.com", 587)
    # server.set_debuglevel(1)#显示编译信息
    # 向mail发送SMTP "ehlo" 命令
    server.ehlo()
    # 启动TLS模式，mail要求
    server.starttls()
    server.login('data.intern@yrbrands.com', 'Password2015')  # 仅smtp服务器需要验证时

    for tmp_custom in range(0,len(customs)-1,1):
        excel_content.append(customs[tmp_custom])
        #deal with the custom
        tmp_shop_name = customs[tmp_custom][title]
        if (customs[tmp_custom][title] == customs[tmp_custom+1][title]) :
            continue
        print u'处理编号为:'+ tmp_custom.__str__()+':'+tmp_shop_name
        print u'finish:'+tmp_custom.__str__()+'/'+ (str) (len(customs)-1)
        emails = ''
        email_send_address = get_the_main_address(customs[tmp_custom][title])#发送人
        emails+=(get_the_copy_addresses(customs[tmp_custom][area]))#抄送人
        # create_package(tmp_shop_name)#创建文件夹
        file_path = create_excel_file(tmp_shop_name,excel_content)#在文件夹中创建excel表格
        outlook(len(excel_content),email_send_address,emails,tmp_shop_name)#发邮件
        # stmp(server,len(excel_content),email_send_address, emails, tmp_shop_name, file_path)
        num = num + len(excel_content)
        if(customs[tmp_custom] != customs[tmp_custom+1]):
            excel_content = []  # clean up
    if(len(excel_content) != 0):
        emails = ''
        excel_content.append(customs[tmp_custom+1])
        email_send_address = get_the_main_address(customs[tmp_custom][title])  # 发送人
        emails+= get_the_copy_addresses(customs[tmp_custom][area])  # 抄送人
        # create_package(tmp_shop_name)  # 创建文件夹
        file_path = create_excel_file(tmp_shop_name, excel_content)  # 在文件夹中创建excel表格
        outlook(len(excel_content),len(excel_content),emails,tmp_shop_name)#发邮件
        # stmp(server,len(excel_content),email_send_address,emails,tmp_shop_name,file_path)
        num = num + len(excel_content)

    server.quit()
    print num
    # 用smtp发送邮件
