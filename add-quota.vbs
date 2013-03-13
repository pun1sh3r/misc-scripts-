add_attribs

Sub add_attribs()
Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = "<LDAP://OU=Roaming-profile-project,OU=Bailiffs Unrestricted,OU=AOC Users,DC=D0300002,DC=AOC>;" & _
  "(&(objectclass=user)(objectcategory=person));" & _
  "adspath,distinguishedname,sAMAccountName;subtree"
	Set objRecordSet = objCommand.Execute
	Do Until objRecordSet.EOF

  	Set objuser = GetObject(objRecordSet.Fields(0).Value)
  	
    	'strpath = "\\S0300100\profiles\" & objRecordset.Fields("samAccountName")
    
   	'WScript.Echo strpath
    
    	Set colDiskQuotas = CreateObject("Microsoft.DiskQuota.1")

	colDiskQuotas.Initialize "E:\", True
	Set objUser = colDiskQuotas.AddUser(objRecordset.Fields("samAccountName"))
	Set objUser = colDiskQuotas.FindUser(objRecordset.Fields("samAccountName"))
	objUser.QuotaLimit = 10000000000	
	objUser.QuotaThreshold = 900000000
  	  
  
  
  objRecordSet.MoveNext

Loop


End Sub 
