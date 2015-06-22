import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application
from os import walk
import os
import sys



_OFFICE365_user = sys.argv[1]
_OFFICE365_pwd = sys.argv[2]
_SEND_FOLDER = 'send'
_RECEIVER = sys.argv[3]
_DOMAIN = _OFFICE365_user.split('@')[1]
_MAIL_SERVER = "smtp.office365.com"

def find_files_in_folder():
	f = []
	for (dirpath, dirnames, filenames) in walk(_SEND_FOLDER):
		f.extend(filenames)
		break
	return f


def send_message(receiver, subject, body, att_file):
	msg = MIMEMultipart('alternative')
	to = receiver

	msg["Subject"] = subject
	msg["From"] = _OFFICE365_user
	content = MIMEText(body,'plain')
	msg.attach(content)
	filename = att_file
	fp = open('%s/%s' %(_SEND_FOLDER, filename),'rb')
	att = email.mime.application.MIMEApplication(fp.read())
	fp.close()
	att.add_header('Content-Disposition','attachment',filename=filename)
	msg.attach(att)

	smtpserver = smtplib.SMTP(_MAIL_SERVER,587)
	smtpserver.ehlo(_DOMAIN)
	smtpserver.starttls()
	smtpserver.ehlo(_DOMAIN)
	smtpserver.login(_OFFICE365_user,_OFFICE365_pwd)
	smtpserver.sendmail(_OFFICE365_user, to, msg.as_string())
	smtpserver.close()


for s_file in find_files_in_folder():
	print s_file
	try:
		send_message(_RECEIVER, 'Test 2', 'Ok its working', s_file)
	except Exception, e:
		print "Error! %s" %e
	else: 
		os.remove("%s/%s" %(_SEND_FOLDER, s_file))
		print "Done!"


