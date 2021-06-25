from email.mime.text import MIMEText
from email.utils import formatdate
import smtplib

def sending_mail(from_addr, to_addr, subject, body):
	msg = MIMEText(body)
	msg['Subject'] = subject
	msg['From'] = from_addr
	msg['To'] = to_addr
	msg['Date'] = formatdate()

	smtpobj = smtplib.SMTP('xxxx.xxxxx.com', 999)
	smtpobj.ehlo()
	smtpobj.starttls()
	smtpobj.ehlo()
	smtpobj.login("xxxx@gmail.com", "password")
	smtpobj.sendmail(from_addr, to_addr, msg.as_string())
	smtpobj.close()
