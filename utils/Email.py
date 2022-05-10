
import os
import os.path

import base64
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

import smtplib

class MAIL():

    def __init__(self,to, file):

        self.user     = 'luis.garcia@guatevirtualmall.com'
        self.password = 'Pass@123'
        self.to       = to
        self.fromx    = 'luis.garcia@guatevirtualmall.com'
        self.file     = file


    def send(self):
        sent_from = self.user
        to = self.to
        subject = 'Generated Excel  Files'
        body = 'Hi!,  this'

        msg = MIMEMultipart()
        msg['Subject'] = 'Generated Excel'

        msg['From'] = 'luis.garcia@guatevirtualmall.com'
        msg['To']   = 'fernandoforce@gmail.com'
        msg.preamble = 'excel '


        filename = self.file

        fp = open(filename, 'rb')
        xls = MIMEBase('application','vnd.ms-excel')
        xls.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(xls)
        xls.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(xls)



        email_text = """\
        From: %s
        To: %s
        Subject: %s

        %s
        """ % (sent_from, ", ".join(to), subject, body)

        try:
            smtp_server = smtplib.SMTP_SSL('mail.guatevirtualmall.com', 465)
            smtp_server.ehlo()
            smtp_server.login(self.user, self.password)
            smtp_server.sendmail(sent_from, to, msg.as_string())
            smtp_server.close()
            print ("Email sent successfully!")
        except Exception as ex:
            print ("Something went wrong….",ex)

