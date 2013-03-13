import smtplib
import os
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders

def send_mail(send_from, send_to, subject, text, file, server):
  
 

  msg = MIMEMultipart()
  msg['From'] = send_from
  msg['To'] = COMMASPACE.join(send_to)
  msg['Date'] = formatdate(localtime=True)
  msg['Subject'] = subject

  msg.attach( MIMEText(text) )

  
  part = MIMEBase('application', "octet-stream")
  part.set_payload( open(file,"rb").read() )
  Encoders.encode_base64(part)
  part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(file))
  msg.attach(part)

  smtp = smtplib.SMTP(server)
  smtp.sendmail(send_from, send_to, msg.as_string())
  smtp.close()



send_mail('nina@test.com','test@test.com', 'please see report attached','report is ready', 'report.pdf', '127.0.0.1')
