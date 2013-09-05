#!/usr/bin/python
import gdata.service
import gdata.spreadsheet
import gdata.spreadsheet.service
import gdata.spreadsheet.text_db
import pprint
import string
from pyad import *
import datetime
import smtplib
from random import choice
import random



client = gdata.spreadsheet.service.SpreadsheetsService()
key = "document-key"
client.email = 'test@test.com'
client.password = 'password'
client.ProgrammaticLogin()
feed1 = client.GetWorksheetsFeed(key)
id_parts1 = feed1.entry[0].id.text.split('/')
curr_wksht_id = id_parts1[len(id_parts1) - 1]
feed = client.GetListFeed(key)
log_file = open("logFile.log","w")
now = datetime.datetime.now()
confirm = ''
mailbod = "mail-body"




def sendEmail(body=None,sender=None,receiver=None):
	try:
		smtpObj = smtplib.SMTP("1.1.1.1")
		smtpObj.sendmail(sender,receiver,body)
	except smtplib.SMTPException as e:
		log_file.write(str(now) + str(e))	
	
def addUser(name,username,email):
	#print email
	confirmation = ''
	pyad.set_defaults(ldapserver="dc.local", username="docbot",password="password")
	ou = adcontainer.ADContainer.from_dn("OU=Users,DC=dc,DC=local")
	passwd_gen = lambda length, ascii =  string.ascii_letters + string.digits + string.punctuation: "".join([list(set(ascii))[random.randint(0,len(list(set(ascii)))-1)] for i in range(length)])
	passwd = passwd_gen(20)
	mailbody = "%s %s %s" % (email,username,passwd)
	try:
		cuser = ou.create_user(username,password=passwd,upn_suffix=None,enable=True,optional_attributes={"description" : "penLabUser","displayName": name, "mail" : email})
		confirmation = "user %s has been created\n" % (username)
		sendEmail(mailbody,"noreply@test.net","test.net")
	except:
		log_file.write(str(now) + " *****error user %s already exist****\n" % (username))
	return confirmation
	
for i,entry in enumerate(feed.entry):
	i += 2
	aduser = entry.custom['ad-user'].text
	name = entry.custom['name'].text
	username = entry.custom['username'].text
	email = entry.custom['email'].text
	if aduser == "FALSE":
		log_file.write(str(now) + " name:%s username:%s email:%s aduser:%s\n" % (name,username,email,aduser))
		client.UpdateCell(row=i,col=10,inputValue="FALSE",key=key,wksht_id=curr_wksht_id)
		confirm += addUser(name,username,email)
sendEmail(mailbod + confirm,"test@test.com","test@test.com")	
	
	
