#-*- coding: utf8 -*-
import poplib
import cStringIO
import email
import base64
#pop3 get email
M=poplib.POP3_SSL('outlook.office365.com')
M.user('data.intern@yrbrands.com')
M.pass_('Password2015')
#number of emails
numMessages=len(M.list()[1])
print 'num of messages',numMessages
for i in range(numMessages):
    m = M.retr(i+1)
    buf = cStringIO.StringIO()
    for j in m[1]:
        print >>buf,j
    buf.seek(0)
    #
    msg = email.message_from_file(buf)
    for part in msg.walk():
        contenttype = part.get_content_type()
        filename = part.get_filename()

        if filename and contenttype=='application/octet-stream':
            #save
            f = open("mail%d.%s.attach" % (i+1,filename),'wb')
            f.write(base64.decodestring(part.get_payload()))
            f.close()