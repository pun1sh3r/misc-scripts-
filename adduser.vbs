

Dim RootDSE,mydomain
Dim input
' strTitle="New User Manager"

set RootDSE=GetObject("LDAP://RootDSE")
'strOU = InputBox("please enter the OU the user belongs",strTitle,"CA_project_Test")
'Set mydomain=GetObject("LDAP://OU=AOC Users,dc=D0300002,dc=AOC")
Set mydomain=GetObject("LDAP://OU=Scripting_Lab,dc=D0300002,dc=AOC")

'set mydomain=GetObject("LDAP://OU=Scripting_lab," & "OU=" & strOU & ",dc=,dc=AOC")
myDomainADSPath=mydomain.ADSPath  	
myDomainPath=MID(mydomain.ADSPath,8)



'move_to_OU "ttester", "County Judge"
'add_attribs
'strcomma = Unescape("%2c")
'WScript.Echo strcomma
'add_to_group "ttester10" , "gato"
get_info
'menu
'get_info_from_file
'WScript.Echo "done"

Sub move_to_OU(strUser, strDesc)
strField = "AdsPath"
strUserDN = connect_query(strField,strUser)
	If InStr(strDesc, "Judge") Or InStr(strDesc, "General Magistrate") >= 1 Then
		Set objNewOU = GetObject("LDAP://OU=Judges,OU=Judicial,OU=AOC Users,DC=D0300002,DC=AOC")
		objNewOU.MoveHere strUserDN, vbNullString
		WScript.Echo strUserDN
		
	ElseIf InStr(strDesc, "General Magistrate") >= 1 Then
		Set objNewOU = GetObject("LDAP://OU=General Magistrates,OU=Judicial,OU=AOC Users,DC=D0300002,DC=AOC")
		objNewOU.MoveHere strUserDN, vbNullString
		
	else
		Set objNewOU = GetObject("LDAP://OU=Scripting_Lab,dc=D0300002,dc=AOC")
		objNewOU.MoveHere strUserDN, vbNullString
		WScript.Echo strUserDN
	End If 
End Sub 


Sub add_attribs()
Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = "<LDAP://OU=CITeS,OU=AOC Users,DC=D0300002,DC=AOC>;" & _
  "(&(objectclass=user)(objectcategory=person));" & _
  "adspath,distinguishedname,sAMAccountName;subtree"
	Set objRecordSet = objCommand.Execute
	Do Until objRecordSet.EOF

  	Set objuser = GetObject(objRecordSet.Fields(0).Value)
  	
  	strAddREG = "Richard E. Gerstein Justice Building" & vbCrlf & "1351 NW 12 Street"
  	strAddJJC = "Juvenile Justice Center" & vbCrLf &  "3300 N. W. 27th Avenue"
  	strAddDCC = "Dade County Courthouse" & vbCrLf &  "73 West Flagler Street"
  	
  	
    'strOffice = objuser.Get("physicalDeliveryOfficeName")
    'WScript.Echo objuser.Get("streetAddress")
    'objuser.put "title", objuser.Get("description")
    'objuser.Put "l" , "Miami"
    'objuser.Put "st", "FL"
    'objuser.Put "c" ,"US"
    'If InStr(strOffice, "REG") >= 1 Then
    'objuser.put "company", "Criminal Division"
    'objuser.Put "streetAddress" , strAddREG
    'objuser.Put "postalCode" , "33125"
    'WScript.Echo "done"
    
    'ElseIf InStr(strOffice, "JJC") >= 1 Then
    ' objuser.put "company", "Juvenile Division"
    'objuser.Put "streetAddress" , strAddJJC
    'objuser.Put "postalCode" , "33142"
    
    'ElseIf InStr(strOffice, "DCC") >= 1 Then
    ' objuser.put "company", "Civil Division"
    'objuser.Put "streetAddress" , strAddDCC
    'objuser.Put "postalCode" , "33130"
    'End If
    strpath = "\\S030005\profiles\" & objRecordset.Fields("samAccountName")
    
    WScript.Echo strpath
    'objuser.Put "profilePath", strpath
    
  	'objuser.SetInfo
  
  
  
  objRecordSet.MoveNext

Loop


End Sub 




Function connect_query(strField, strSAM)
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"	
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = "<LDAP://dc=D0300002,dc=AOC>;" & "(&(objectCategory=person) (objectClass=user)" & "(sAMAccountName=" & strSAM &"));" & "AdsPath, DistinguishedName;subtree"
	Set objRecordSet = objCommand.Execute
	
	connect_query = objRecordSet.Fields(strField)
	
	
	
End function



Sub add_to_group(strSAM, strDesc)
	strField = "AdsPath"
	Set objGroup= GetObject("LDAP://OU=AOC Groups,dc=D0300002,dc=AOC")
	'Set objGroup= GetObject("LDAP://OU=Scripting_Lab,dc=D0300002,dc=AOC")
	Set objConnection = CreateObject("ADODB.Connection")
	 
	If InStr(strDesc, "Judge") >=1 Then 	
		objGroup.Filter = Array("group")
		For Each group In objGroup
	
			If InStr(group.name, "Judges Group")   Then
				group.add(connect_query(strField, strSAM))
				WScript.Echo " jd group 1 exist" 
				
				
			ElseIf InStr(group.name, "eBench Courtroom Users") Then
				group.add(connect_query(strField, strSAM))
				WScript.Echo " jd group 2 exist"  
				
			End If 
		
		Next
	ElseIf InStr(strDesc, "Judicial Assistant")Or InStr(strDesc, "Administrative Secretary") >= 1 Then 
		objGroup.Filter = Array("group")
		For Each group In objGroup
			If StrComp(group.name ,"Judicial Assistants") = 0 Then 
				group.add(connect_query(strField, strSAM))
				WScript.Echo "JA group exist"
			
			End If 
		
		Next
	ElseIf InStr(strDesc, "Bailiff") >= 1 Then 
		objGroup.Filter = Array("group")
		For Each group In objGroup
			If InStr(group.name, "Bailiffs") Then 
				group.add(connect_query(strField, strSAM))
				WScript.Echo "Bailiff group exist"
			
			End If 
		
		Next
	
	Else 
		objGroup.Filter = Array("group")
		For Each group In objGroup
			If InStr(group.name, "All AOC Staff") Then 
				group.add(connect_query(strField, strSAM))
				WScript.Echo "aoc group exist"
			
			End If 
		
		Next
	End If 
End Sub 



Sub get_info_from_file()
	'Dim samArray(2)
	Dim	strdata
	Dim strname
	Dim strFname
	Dim strLname
	Dim strSAM
	Dim strDesc
	Dim strOffice
	Dim strPhone
	Dim intComparator
	Dim IntComparator1
	Dim intCOmparator2
	Const ForReading = 1
	Set objFSO = CreateObject("Scripting.FilesystemObject")
	Set objTextFile = objFSO.OpenTextFile("C:\Documents and Settings\lmendieta\Desktop\bailiff-project\test1.csv", ForReading)
		Do until objTextFile.AtEndOfStream
			strdata = objTextFile.ReadLine
			'strcomma = Chr(34)
			strcomma = Unescape("%2c")
			'WScript.Echo strdata
			arrayList = Split(strdata , ",")
			strname = arrayList(0) & " "  & arrayList(1)
			strLname = arrayList(0)
			strFname = arrayList(1)
			
			strDesc = "bailiff for " & " " & arrayList(2)
			'strDesc = "test" & " " & "bailiff for " & " " & arrayList(2)& " " & arrayList(3)
			strOffice = arrayList(4)
			'strOffice = arrayList(5)& "-" & arrayList(4)
			strPhone =  arrayList(3)
			'strPhone = "(305)-" & arrayList(6)
			strSAM = LCase(Left(strFname,1)) & LCase(strLname)
			strSAM1 = LCase(Left(strFname,2)) & LCase(strLname)
			strSAM2 = LCase(Left(strFname,3)) & LCase(strLname)
			
			If idExist(strSAM) = True Then	
				strSAM = strSAM1
				If idExist(strSAM1)= True Then
				 strSAM = strSAM2
				End If 
			End If
		
			On Error Resume next
			add_user strname, strSAM, strFname, strLname, strDesc, strOffice, strPhone 	
		Loop
	objTextFile.Close()
End Sub 


Sub menu()
	Dim CA(3)
	CA(0)="CB03"
	CA(1)="CC03"
	CA(2)="CD03"
	CA(3)="CE03"
	Dim email(3)
	email(0)="emergencymotions"
	email(1)="motioncalendar"
	email(2)="proposedorders"
	email(3)="specialsets"
	strserver = "@S0300063.D0300002.AOC"
	strDescString = "Server S0300063 Sharepoint"
	strMenu="Select a Judicial Section" & VbCrLf &_
	 "1 CA 03"  & VbCrLf &_
	 "2 CA 11" & VbCrLf &_
	 "3 CA 00"
	 
	 rc=InputBox(strMenu,"Judicial Section",1)
	 If IsNumeric(rc) Then
	     Select Case rc
	         Case 1
	         	For Each judsec In CA 
	         	 	For Each mail In email
	         	 	create_contact judsec, mail, strserver, strDescString
	         	 	create_ca_user judsec, mail 
	             	Next
	             Next
	         Case 2
	             WScript.Echo "CA11"
	         Case 3
	             WScript.Echo "CA00"
	         Case Else                                                   
	             wscript.echo "That's not on the menu."
	     End Select
	 End if
End sub

Sub create_ca_user(strCA, strmail)
	set mydomain1=GetObject("LDAP://OU=service Accounts,OU=CA_Test_OU,OU=Scripting_Lab,dc=D0300002,dc=AOC")
	myDomainADSPath=mydomain1.ADSPath		
	myDomainPath=MID(mydomain1.ADSPath,8)
	Set objContainer=GetObject("LDAP://" & myDomainPath)
	Set objUser=objContainer.Create("user","CN=" & StrCA & strmail)
	Dim strMailDN
	objUser.Put "sAMAccountName", strCA & strmail
	objUser.Put "userPrincipalName", strCA & strmail & "@D0300002.AOC"
	objUser.SetInfo
	enable_acct(objUser)
	set_passwd(objUser)
	add_attrib objUser, strCA & strmail, " blank " ,strCA & strmail
	'strMailDN=SelectMailStore
	'strMailDN = "CN=Judicial Assistants Limited Storage Mail Store  (S0300088),CN=Judicial Mailboxes,CN=InformationStore,CN=S0300088,CN=Servers,CN=JUD11,CN=Administrative Groups,CN=FLCOURTS,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=D0300002,DC=AOC"
	'strMailDN = "CN=Judges and Magistrates Mail Store  (S0300088),CN=Judicial Mailboxes,CN=InformationStore,CN=S0300088,CN=Servers,CN=JUD11,CN=Administrative Groups,CN=FLCOURTS,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=D0300002,DC=AOC"
	strMailDN = "CN=AOC Limited Storage Mail Store  (S0300088),CN=AOC Mailboxes,CN=InformationStore,CN=S0300088,CN=Servers,CN=JUD11,CN=Administrative Groups,CN=FLCOURTS,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=D0300002,DC=AOC"
	MakeMail objUser,strMailDN
	strForwardingAddress = "CN="& strCA & strmail &",OU=CA_Test_OU,OU=Scripting_Lab,dc=D0300002,dc=AOC" 
	strUserAdsPath = "CN="& strCA & strmail & ",OU=service Accounts,OU=CA_Test_OU,OU=Scripting_Lab,DC=D0300002,DC=AOC" 
	Set objUser1 = GetObject("LDAP://" & strUserAdsPath) 
	objUser1.altRecipient = strForwardingAddress 

	objUser1.SetInfo
	
End Sub 

'create_contact
Sub create_contact(strCA, strcontactName, strserver, strDescString)
	set mydomain=GetObject("LDAP://OU=CA_Test_OU,OU=Scripting_Lab,dc=D0300002,dc=AOC")
	myDomainADSPath=mydomain.ADSPath		
	myDomainPath=MID(mydomain.ADSPath,8)
	Set objContainer=GetObject("LDAP://" & myDomainPath)
	Set objContact=objContainer.Create("contact","CN="&strCA & strcontactName)
	objContact.Put "givenName",strCA & strcontactName
	objContact.Put "displayName", strCA & strcontactName
	objContact.Put "Description", strDescString & " " & strCA & " " & strcontactName & " " & "Document Library"
	objContact.Put "Mail", strCA& strcontactName & strserver
	objContact.Put "proxyAddresses", "SMTP:" & strCA& strcontactName & strserver
	objContact.Put "targetAddress", "SMTP:" & strCA& strcontactName & strserver
	objContact.Put "mailNickname", strCA& strcontactName
	objContact.SetInfo
	'strMailDN=SelectMailStore

	'MakeMail objUser,strMailDN	
End sub

' gather user info for further processing 
Sub get_info()
	strName=InputBox("What is the name of the new user",strTitle,"Lastname ,Firstname")
	strSAM=InputBox("What will the be the user's logon name?",strTitle,"")
	While idExist(strSAM) = True
		strSAM=InputBox("User ID" & " " & strSAM & " " & "is already being used please try again",strTitle,"")
		idExist(strSAM)
	Wend
	strDesc=InputBox("please enter user Position",strTitle)
	strSplit = Split(strName)
	strFname = strSplit(1)
	strLname = strSplit(0)
	strOffice = " "
	strPhone = " "
	add_user strName,strSAM, strFname, strLname, strDesc, strOffice, strPhone

End Sub 

'add user to Active directory 
Sub add_user(strName, strSAM, strFname, strLname, strDesc,strOffice,strPhone)
	Dim strMailDN
	Set objContainer=GetObject("LDAP://" & myDomainPath)
	Set objUser=objContainer.Create("user","CN="&strName)
	objUser.Put "sAMAccountName", strSAM
	objUser.Put "userPrincipalName", strSAM & "@D0300002.AOC"
	objUser.SetInfo
	enable_acct(objUser)
	set_passwd(objUser)
	add_attrib objUser, strFname, strLname, strDesc, strOffice, strPhone
	add_to_group strSAM, strDesc
	'strMailDN=SelectMailStore
	If InStr(strDesc, "Judicial Assistant")Or InStr(strDesc, "Administrative Secretary") >= 1 Then
		'WScript.Echo strDesc & " " & " match"
		strMailDN = "CN=Judicial Assistants Limited Storage  Mail Store (S0300088),CN=Judicial Mailboxes,CN=InformationStore,CN=S0300088,CN=Servers,CN=JUD11,CN=Administrative Groups,CN=FLCOURTS,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=D0300002,DC=AOC"
		
	ElseIf InStr(strDesc, "Senior Judge") Or InStr(strDesc, "County Court Judge") Or InStr(strDesc, "Circuit Court Judge") Or InStr(strDesc, "General Magistrate") Or InStr(strDesc, "Judge") >= 1 Then
		'WScript.Echo strDesc & " " & " match"
		strMailDN = "CN=Judges and Magistrates Mail Store (S0300088),CN=Judicial Mailboxes,CN=InformationStore,CN=S0300088,CN=Servers,CN=JUD11,CN=Administrative Groups,CN=FLCOURTS,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=D0300002,DC=AOC"
	Else 	
		WScript.Echo "no match"
		strMailDN = "CN=AOC Limited Storage Mail Store  (S0300088),CN=AOC Mailboxes,CN=InformationStore,CN=S0300088,CN=Servers,CN=JUD11,CN=Administrative Groups,CN=FLCOURTS,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=D0300002,DC=AOC"	
	End If 
	MakeMail objUser,strMailDN
End Sub


Function idExist(strSAM)
	Dim intFlag 
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = "<LDAP://dc=D0300002,dc=AOC>;" & "(&(objectCategory=person) (objectClass=user)" & "(sAMAccountName=" & strSAM &"));" & "sAMAccountName, DistinguishedName;subtree"
	Set objRecordSet = objCommand.Execute
	If objRecordSet.RecordCount = 0 Then 
		'WScript.Echo strSAM & " " & ": false"
		idExist = False
		Exit Function
	Else
		idExist = true
		'WScript.Echo strSAM & " " & ": true"
		Exit Function
		
	End if
End Function


'set default password for user account
Sub set_passwd(objUser)

	objUser.SetPassword "hellohello"
	objUser.Put "pwdLastSet", 0
	objUser.SetInfo

End Sub

'set user attributes to respective acct
Sub add_attrib(objUser, strFname, strLname, strDesc, strOffice, strPhone)

	objUser.Put "givenName" , strFname
	'objUser.Put "name" , strFname
	objUser.Put "sn", strLname
	objUser.Put "description", strDesc
	objUser.Put "physicalDeliveryOfficeName", strOffice
	objUser.Put "telephoneNumber", strPhone
	objUser.SetInfo
End Sub

'enable user acct
Sub enable_acct(objUser)
Dim intUAC
	Const ADS_UF_ACCOUNTDISABLE = 2
	
	
	intUAC = objUser.Get("userAccountControl")
		If intUAC And ADS_UF_ACCOUNTDISABLE Then
			objUser.Put"userAccountControl", intUAC Xor ADS_UF_ACCOUNTDISABLE
			objUser.SetInfo
		End If
End Sub

'create Mailbox Function
Sub MakeMail(objUser,strMailDN)
	On Error Resume next
	Err.Clear
	'WScript.Echo strMail.DN	
	objUser.CreateMailbox strMailDN
		If Err.Number <>0 Then
			f.WriteLine Now & " Error creating mailbox for " & objUser.ADSPath & " on " & strMailDN 
			f.WriteLine Now & " " & Err.Number & " " & Err.Description
		Else
			'objUser.Put "altRecipient", "CN=ca03emergencymotions,OU=CA Contacts for Sharepoint,OU=Foreign Accounts,DC=D0300002,DC=AOC"
			objUser.SetInfo
			WScript.Echo "Created mailbox for " & objUser.Name
		End If

End Sub



