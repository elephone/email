#coding: utf-8
import smtplib
from email.mime.text import MIMEText
from email.header import Header

sender = 'data.intern@yrbrands.com'
receiver = '574307361@qq.com'
#receiver=' xxxxxx @qq.com'
username = 'data.intern@yrbrands.com'
password = 'Password2015'

subject = 'python email test'
#中文需参数‘utf-8’，单字节字符不需要
msg = MIMEText('你好呀，最近好么','plain','utf-8')
msg['Subject'] = Header(subject, 'utf-8')
msg["From"] = sender
msg["To"] = str(receiver)
msg["Reply-To"] = sender
print msg.as_string()

smtpserver = 'smtp.office365.com'
server_port = 587
# print email_message
#创建SMTP对象
smtp = smtplib.SMTP(smtpserver,server_port)
smtp.set_debuglevel(1)
#向mail发送SMTP "ehlo" 命令
smtp.ehlo()
#启动TLS模式，mail要求
smtp.starttls()
#用户验证
smtp.login(username, password)
#发送邮件
smtp.sendmail(sender, receiver, msg.as_string())
#退出
smtp.quit()