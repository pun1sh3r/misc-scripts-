On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = _
    "SELECT Name FROM 'LDAP://OU=Bailiffs,OU=AOC Users,dc=D0300002,dc=AOC' WHERE objectCategory='user'"  
Set objRecordSet = objCommand.Execute

Wscript.Echo objRecordSet.RecordCount
