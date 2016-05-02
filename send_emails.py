#coding:utf-8

import win32com.client as win32
import datetime
from email.MIMEText import MIMEText
from email.MIMEMultipart import MIMEMultipart
import email.MIMEBase

from email import Utils,Encoders
import mimetypes,sys
from email import Utils,encoders
import os
def outlook(excel_inf_num,to_address='',cc_address=''):
    now = datetime.date.today()
    end_send_time = now + datetime.timedelta(days=3)
    # app= 'Outlook'
    # olook = win32.gencache.EnsureDispatch("%s.Application" % app)

    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = u'{0}_撼路者销售线索跟进'.format(now)
    newMail.Body =u'''各位撼路者经销商同仁 您好，

我们是Wunderman的CRM团队，目前正在以电访的形式对每日新增撼路者销售线索进行甄别，今日获取到的有效销售线索如附件。鉴于线索时效性，请各位务必在以下规定时间内完成跟进，并反馈追踪结果。
{0} 分发{1} 条销售线索，烦请各位及时完成跟进，务必将有效线索录入DMS系统，并将追踪结果于{2} 12点之前反馈到邮箱：data.intern@yrbrands.com
如有任何问题，烦请及时与我们联系，谢谢！

江铃福特
 '''.format(now,excel_inf_num,end_send_time)
    newMail.To = '574307361@qq.com'
    newMail.CC = '574307361@qq.com'+'; huanhuan9312@163.com'

    # newMail.CC = '574307361@qq.com'
    # newMail.CC = '574307361@qq.com'
    # newMail.Save()
    print "send ok"
#===================================
    # # 添加附件就是加上一个MIMEBase，从本地读取一个图片:
    # with open('test.png', 'rb') as f:
    #     # 设置附件的MIME和文件名，这里是png类型:
    #     mime = MIMEBase('image', 'png', filename='test.png')
    #     # 加上必要的头信息:
    #     mime.add_header('Content-Disposition', 'attachment', filename='test.png')
    #     mime.add_header('Content-ID', '<0>')
    #     mime.add_header('X-Attachment-Id', '0')
    #     # 把附件的内容读进来:
    #     mime.set_payload(f.read())
    #     # 用Base64编码:
    #     encoders.encode_7or8bit(mime)
    #     # 添加到MIMEMultipart:

    # ## 设置附件头
    # basename = os.path.basename(file_name)
    # file_msg.add_header('Content-Disposition', 'attachment', filename=basename)  # 修改邮件头
    # main_msg.attach(file_msg)

    # 构造MIMEMultipart对象做为根容器
    main_msg = email.MIMEMultipart.MIMEMultipart()

    # 构造MIMEText对象做为邮件显示内容并附加到根容器
    text_msg = email.MIMEText.MIMEText("我this is a test text to text mime", "utf-8")
    main_msg.attach(text_msg)

    file_name = u'分发销售线索名单426.xlsx'

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
    basename = os.path.basename(file_name)
    file_msg.add_header('Content-Disposition', 'attachment', filename=basename)  # 修改邮件头
    main_msg.attach(file_msg)
    # 得到格式化后的完整文本
    fullText = main_msg.as_string()

    newMail.Save()
    #============================================================================================
    #以下为对收件箱、已发送、发件箱、草稿箱、已删除、任务进行的操作：
    ns = obj.GetNamespace("MAPI")

    #收件箱
    inbox = ns.GetDefaultFolder(win32.constants.olFolderInbox)
    messages1 = inbox.Items
    print u"收件箱邮件数量 :",messages1.Count

    #已发送邮件数量
    sentmail = ns.GetDefaultFolder(win32.constants.olFolderSentMail)
    messages2 = sentmail.Items
    print u"已发送邮件数量 :",messages2.Count

    # #已删除邮件数量
    # DeletedItems = ns.GetDefaultFolder(win32.constants.olFolderDeletedItems)
    # messages3 = DeletedItems.Items
    # print u"已删除邮件数量 :",messages3.Count

    #草稿数量
    drafts = ns.GetDefaultFolder(win32.constants.olFolderDrafts)
    messages4 = drafts.Items
    messages41 = drafts.Items

    print u"草稿邮件数量 :",messages4.Count

    #发件箱
    outbox = ns.GetDefaultFolder(win32.constants.olFolderOutbox)
    messages5 = outbox.Items
    print u"发件箱邮件数量 :",messages5.Count
    # obox.Display() #打开发件箱
    # obox.Items.Item(1).Display()#打开发件箱中第一封邮件

    #任务操作
    task_list = ns.GetDefaultFolder(win32.constants.olFolderTasks)
    tasks = task_list.Items
    print u"任务数量:", tasks.Count

import win32com.client


def send_mail_via_com(text='try', subject='try', recipient='574307361@qq.com', profilename="Outlook2010"):
    s = win32.Dispatch("Mapi.Session")
    o = win32.Dispatch("Outlook.Application")
    s.Logon(profilename)

    Msg = o.CreateItem(0)
    Msg.To = recipient

    # Msg.CC = "moreaddresses here"
    # Msg.BCC = "address"

    Msg.Subject = subject
    Msg.Body = text

    attachment1 = "Path to attachment no. 1"
    attachment2 = "Path to attachment no. 2"
    Msg.Attachments.Add(attachment1)
    Msg.Attachments.Add(attachment2)

    Msg.Send()


if __name__ =="__main__":
    # send_mail_via_com()
    outlook(1)