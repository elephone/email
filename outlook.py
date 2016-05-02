#coding:utf-8

import win32com.client as win32
RANGE = range(3,8)

def outlook():
    app= 'Outlook'
    olook = win32.gencache.EnsureDispatch("%s.Application" % app)
    mail=olook.CreateItem(win32.constants.olMailItem)
    mail.Recipients.Add('jianl@163.com')
    mail.Recipients.Add('jianliy@sina.com')
    subj = mail.Subject = '21112Python -to - %s ' %app
    body = ["line %d" % i for i in RANGE]
    # body.insert(0,"%s\r\n" %subj)
    body.append("\r\nThat's all folks!")
    mail.Body = '\r\n'.join(body)
    mail.Send()
    print "send ok"
    #============================================================================================
    #以下为对收件箱、已发送、发件箱、草稿箱、已删除、任务进行的操作：
    ns = olook.GetNamespace("MAPI")

    #收件箱
    inbox = ns.GetDefaultFolder(win32.constants.olFolderInbox)
    messages1 = inbox.Items
    print u"收件箱邮件数量 :",messages1.Count

    #已发送邮件数量
    sentmail = ns.GetDefaultFolder(win32.constants.olFolderSentMail)
    messages2 = sentmail.Items
    print u"已发送邮件数量 :",messages2.Count

    #已删除邮件数量
    DeletedItems = ns.GetDefaultFolder(win32.constants.olFolderDeletedItems)
    messages3 = DeletedItems.Items
    print u"已删除邮件数量 :",messages3.Count

    #草稿数量
    drafts = ns.GetDefaultFolder(win32.constants.olFolderDrafts)
    messages4 = drafts.Items
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

if __name__ =="__main__":
    outlook()