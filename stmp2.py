# -*- coding: utf-8 -*-
import smtplib
import email.MIMEMultipart  # import MIMEMultipart
import email.MIMEText  # import MIMEText
import email.MIMEBase  # import MIMEBase
import os.path

import sys
reload(sys)
sys.setdefaultencoding('utf8')

From = "data.intern@yrbrands.com"
To = "574307361@qq.com"
file_name = ur'分发销售线索名单427.xlsx'  # 附件名



# 构造MIMEMultipart对象做为根容器
main_msg = email.MIMEMultipart.MIMEMultipart()

# 构造MIMEText对象做为邮件显示内容并附加到根容器
text_msg = email.MIMEText.MIMEText("我this is a test text to text mime","utf-8")
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
basename = os.path.basename(file_name).encode("utf-8")
file_msg.add_header('Content-Disposition', 'attachment', filename=basename)  # 修改邮件头
main_msg.attach(file_msg)

# 设置根容器属性
main_msg['From'] = From
main_msg['To'] = To
main_msg['Subject'] = "attach test "
main_msg['Date'] = email.Utils.formatdate()

# 得到格式化后的完整文本
fullText = main_msg.as_string()


server = smtplib.SMTP("smtp.outlook.com",587)
server.set_debuglevel(1)
#向mail发送SMTP "ehlo" 命令
server.ehlo()
#启动TLS模式，mail要求
server.starttls()
server.login('data.intern@yrbrands.com', 'Password2015')  # 仅smtp服务器需要验证时
# 用smtp发送邮件
try:
    server.sendmail(From, To, fullText)
except smtplib.SMTPException,ex:
    print smtplib.SMTPException,ex
finally:

    server.quit()
