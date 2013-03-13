#!/usr/bin/env python

import pexpect
import sys 
import optparse
usr = ""
passwd = ""
enable = ""
server = ""

cont = "n"
ip_addrs = []




def restore_config(usr,ip,passwd,hostname,config_file):
  child = pexpect.spawn('telnet %s' % (ip))
	child.expect(['sername:',pexpect.EOF, pexpect.TIMEOUT])	
	child.sendline(passwd)
	child.expect(['assword:',pexpect.EOF, pexpect.TIMEOUT])
	child.sendline(passwd)
	child.sendline("enable")
	child.expect(['assword:',pexpect.EOF, pexpect.TIMEOUT])
	child.sendline(passwd)
	child.sendline("copy tftp: running-config")
	child.expect(['emote ',pexpect.EOF, pexpect.TIMEOUT])
	child.sendline(server)
	child.expect(['ilename ',pexpect.EOF, pexpect.TIMEOUT])
	child.sendline(config_file)
	child.expect(['ilename ',pexpect.EOF, pexpect.TIMEOUT])
	child.sendline("running-config")
	child.sendline("hostname",hostname,child)
	child.interact()	

def addLogging(usr,addr,paswd):
	child1 = ssh_connect(usr,addr,passwd)
	
	child1 = sendline("conf t",child)
	child1 = sendline("no snmp-server community chapulin RO")
	child1 = sendline("snmp-server community chapulin RO 10  ")
	child1 = sendline("logging host 10.3.7.116 ")
	child1 = sendline("logging trap 5 ")
	child1 = sendline("service timestamps log datetime localtime")
	child1 = sendline("no logging console")
	child1 = sendline("no logging monitor")
	child1 = sendline("snmp-server host 10.3.7.206 wth")
	child1 = sendline("snmp-server enable traps snmp authentication linkdown linkup coldstart warmstart")
	child1 = sendline("snmp-server enable traps flash insertion removal")
	child1 = sendline("snmp-server enable traps port-security")
	child1 = sendline("snmp-server enable traps envmon fan shutdown supply temperature status")
	child1 = sendline("snmp-server enable traps license")
	child1 = sendline("snmp-server enable traps config")
	child1 = sendline("snmp-server enable traps errdisable")
	child1 = sendline("snmp-server enable traps tty")
	child1 = sendline("snmp-server enable traps vlancreate")
	child1 = sendline("snmp-server enable traps vlandelete")
	child1 = sendline("archive")
	child1 = sendline("log config")
	child1 = sendline("logging enable")
	child1 = sendline("logging size 200")
	child1 = sendline("hidekeys")
	child1 = sendline("notify syslog contenttype xml")
	child1 = sendline("exit",)
	child1 = sendline("exit",)
	child1 = sendline("login block-for 60 attempts 3 within 10")
	child1 = sendline("login on-failure log")
	child1 = sendline("login on-success log")
	child1 = sendline("logging source-interface Vlan1")
	child1 = sendline("no access-list 10 permit 172.16.18.0 0.0.0.255")
	child1 = sendline("no access-list 10 permit 10.3.7.0 0.0.0.255 ")
	child1 = sendline("access-list 10 permit 172.16.18.0 0.0.0.255 log")									
	child1 = sendline("do write")
	child1 = sendline("exit")
	child1 = sendline("exit")
	
	child1.interact()
	

def upgradeIOS(usr,addr,passwd,server,ios):
	child = interact(usr,addr,passwd)
	child1 = sendline("dir",child)
	child1 = sendline("delete /force /recursive flash:c3750-ipbasek9-mz.122-40.SE")
	child1 = sendline("archive download-sw /overwrite tftp://172.16.18.71/c3750-ipbasek9-tar.122-53.SE2.tar")
	child1 = sendline("conf t")
	child1 = sendline("boot system flash:/c3750-ipbasek9-mz.122-53.SE2/c3750-ipbasek9-mz.122-53.SE2.bin")
	child1 = sendline("do write")
	child1 = sendline("do reload")
	child1.interact()

def interact_(usr,addr,passwd,parameter):
	child = pexpect.spawn('ssh %s@%s' % (usr,addr))
	i = child.expect(['assword:',r"yes/no"],timeout=30 )
	if (i == 0):
		child.sendline(passwd)
		
	elif i==1:
		child.sendline("yes")
		child.expect(['assword:'], timeout=30)
		child.sendline(passwd)
		
	child.sendline("enable")
	child.expect(['assword:'], timeout=30)
	child.sendline("")
	child.sendline("conf t ")
	child = send_command("hostname", parameter, child)
	child.sendline("no username cisco")
	child.sendline("int Dot11Radio0")
	child.sendline("no shut")	
	child.interact()
def chpasswd(usr,addr,passwd,parameter ):
	child = ssh_connect(usr,addr,passwd)
	child1 = sendline("enable")
	child1 = sendline("conf t")
	child1 = send_command("no username ",parameter ,child)
	child1 = send_command("username  privilege 15 secret ",parameter, child)
	child1 = sendline("enable secret ",child)
	child1 = sendline("no snmp-server community public RO")
	child1 = sendline("snmp-server community public RO")
	child1 = sendline("exit",child)
	child1 = sendline("wr mem",child)
	child1 = sendline("exit",child)
	child1 = sendline("exit",child)
	child1.interact()

	

def connect(usr,addr,passwd,option):
	try:
		if(option == '1'):
			interact(usr,addr,passwd)
		elif(option == '2'):
			chpasswd(usr,addr,passwd)
		elif(option == '3'):
			upgradeIOS(usr,addr,passwd)
		elif(option == '4'):
			addLogging(usr,addr,passwd)
		#elif(option == '4'):
			#ssh_connect(usr,addr,passwd)

				
	except:
		return 1

def ssh_connect(usr,addr,passwd):
	child = pexpect.spawn('ssh %s@%s' % (usr,addr))
	i = child.expect(['assword:',r"yes/no"],timeout=30 )
	if (i == 0):
		child.expect(['assword:',pexpect.EOF, pexpect.TIMEOUT])	
		child.sendline(passwd)
		sendline("enable")
		child.expect(['assword:',pexpect.EOF, pexpect.TIMEOUT])
		
	elif (i==1):
		child.sendline("yes")	
		
	return child

def send_command(cmd,paramater,child):
	command = cmd + " " + parameter
	child.sendline(command)
	
	return child

def read_ip_file(name):
	ip_file = open(name)
	
	for line in ip_file:
		ip_addrs.append(line)	
		
	ip_file.close()
	return ip_addrs


#option = raw_input("press 1 to connect to switch\npress 2 to change sw password\npress 3 update IOS\noption selected:")
#addr= raw_input("enter device ip address:")
#addr = sys.argv[1]
#connect(usr,addr,passwd,'4')

ip = sys.argv[1]
parameter = sys.argv[2]
interact_(usr,ip,passwd,parameter)

#restore_config(usr,ip,passwd,hostname)


#array = read_ip_file(name)
#for ip in array:
#	connect(usr,ip,passwd,'4')
#print "done!"

##if (option == '1'):
#	addr= raw_input("enter device ip address:")
#	connect(usr,addr,passwd,option)
#elif(option == '2'):
#	array = read_ip_file()
#	for ip in array:
#		connect(usr,ip,passwd,option)
#	print "done!"
	
#elif(option == '3'):
#	addr= raw_input("enter device ip address:")
#	connect(usr,addr,passwd,option)
	 
#else:
#	print "exit"




#connect(usr,addr,passwd)
#cont = raw_input("continue? (y/n):")
#while(cont == "y"):	
#	addr = raw_input("enter ip address:")
#	connect(usr,addr,passwd)
#	cont = raw_input("continue with another switch? (y/n):")


#read_ip_file()
	
