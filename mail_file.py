import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application
from os import walk
import os
import sys
from optparse import OptionParser

_VERSION = "%prog 1.1"


def _handle_opts():
	parser = OptionParser(usage='usage: python %prog [option]',version=_VERSION,
		description='This is Description for %prog' 
			)
	parser.add_option('-t', '--to', action ='store', dest='rec_email', metavar ="EMAIL",
		help= 'Receiver email address.')
	parser.add_option('-f', '--from', action ='store', dest='from_email', metavar="EMAIL",
		help= 'Sender email address.')
	parser.add_option('-p', '--password', action ='store', dest='password', metavar="PASSWORD",
		help= 'Sender password.')
	parser.add_option('-s', '--subject', action ='store', dest='subject', metavar="SUBJECT",default=" ",
		help= 'Subject of email. (optional)')
	parser.add_option('-b', '--body', action ='store', dest='body', metavar="TEXT",default=" ",
		help= 'Body of message. (optional)')
	parser.add_option('-d', '--dir', action ='store', dest='dir', metavar="PATH",
		help= 'Path for sending files.')

	return parser

def _check_opts(options,parser):
	for key,value in options.__dict__.items():
		if not value:
			parser.error("All mandatory options are necessary! use -h for help!")


def find_files_in_folder():
	f = []
	for (dirpath, dirnames, filenames) in walk(_SEND_FOLDER):
		f.extend(filenames)
		break
	return f


def send_message(receiver, body, att_file):
	msg = MIMEMultipart('alternative')
	to = receiver

	msg["Subject"] = _SUBJECT
	msg["From"] = _OFFICE365_user
	msg["To"] = to
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
	smtpserver.sendmail(_OFFICE365_user, to.split(","), msg.as_string())
	smtpserver.close()


parser = _handle_opts()
options,args = parser.parse_args()
_check_opts(options,parser)
if args:
	parser.print_help()
else:
	_OFFICE365_user = options.from_email
	_OFFICE365_pwd = options.password
	_SEND_FOLDER = options.dir
	_RECEIVER = options.rec_email
	_DOMAIN = _OFFICE365_user.split('@')[1]
	_MAIL_SERVER = "smtp.office365.com"
	_SUBJECT = options.subject
	_BODY = options.body
	for s_file in find_files_in_folder():
		print "%s -> %s " %(s_file,_RECEIVER)
		try:
			send_message(_RECEIVER, _BODY, s_file)
		except Exception, e:
			print "Error! %s" %e
		else: 
			os.remove("%s/%s" %(_SEND_FOLDER, s_file))
			print "Done!"


