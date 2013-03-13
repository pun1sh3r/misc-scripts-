On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
strName = "file.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFile = objFSO.CreateTextFile("pcList.txt")
Set objTextFile = objFSO.OpenTextFile(strName, 2, True)


objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE

objCommand.CommandText = _
"SELECT name, description, office  FROM 'LDAP://OU=AOC Users,dc=D0300002,dc=AOC' WHERE objectCategory='user' " 

Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
Wscript.Echo objRecordSet.Fields("Office").Value'objRecordSet.Fields("Name").Value  objRecordSet.Fields("Description").Value
objTextFile.WriteLine objRecordSet.Fields("Name").Value '& objRecordSet.Fields("Description").Value

PCVar = objRecordSet.Fields("Name").Value
objRecordSet.MoveNext
Loop
objTextFile.Close

IF PCVar = "" then
WScript.Echo "No Computer of that name was found in the directory"
END IF
