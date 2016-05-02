# -*- conding:cp936 -*-
import string,time,os
import email
from email.Header import Header
from email.Header import decode_header
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email import Encoders
import base64

def showMessage(mail):
    if mail.is_multipart():
        for part in mail.get_payload():
            showMessage(part)
    else:
        types=mail.get_content_type()
        if types=='text/plain':
            try:
               stmp=mail.get_payload(decode=True)
               print stmp
            except:
               print '[*001*]BLANK'

        elif types=='text/base64':
            try:
               stmp=base64.decodestring(mail.get_payload())
               print stmp
            except:
               print '[*001*]BLANK'



def showMessage4(mail):
    for part in mail.walk():
        contenttype=part.get_content_type()
        if contenttype=='text/plain':
            try:
               print part.get_payload(decode=True)
            except:
               print ''
        elif contenttype=='text/base64':
            try:
                print base64.decodestring(part.get_payload())
            except:
                print ''

def PareMessageAttachMent(mail):
   global PATH
   count=0
   ErrorList=[]

   for part in mail.walk():
      contenttype = part.get_content_type()
      filename = part.get_filename()
      if filename==None:
          continue
      filename=email.Header.decode_header(filename)[0][0]
      filename=os.path.split(filename)[1]                          #防止文件名出现全路径错误

      if filename:                                                 ###解析邮件的附件内容
         fn=PATH+filename
         if os.path.exists(fn)==False:
            f=open(fn,'wb')
            try:
               f.write(base64.decodestring(part.get_payload()))
            except:
               pass
            f.close()
         else:
            continue


PATH='mail/Attach/'

files=os.listdir('mail')
id=0
for eachfile in files:
  if os.path.splitext(eachfile)[1].lower()=='.eml':
    mail=email.message_from_file(open('mail/'+eachfile))
    mail['subject'],mail.get('subject')
    mail['From'],mail.get('From')
    mail['To'],mail.get('To')
    mail['date'],mail.get('date')
    subject = email.Header.decode_header(mail['subject'])[0][0]
    subcode=  email.Header.decode_header(mail['subject'])[0][1]
    FromAddr=email.Header.decode_header(mail['From'])[0][0]
    ToAddr=email.Header.decode_header(mail['To'])[0][0]
    date=email.Header.decode_header(mail['date'])[0][0]

    print '   ********* mail ',id+1,' *********'
    id=id+1

    print eachfile
    showMessage(mail)
    PareMessageAttachMent(mail)
    print ' '

