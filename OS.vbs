Sub checkOs()


  strComputer = "."
	 Dim objWMI, objItem, colItems
	 Dim OSVersion, OSName, ProductType
	 Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	Set colOperatingSystems = objWMIService.ExecQuery _
	    ("Select * from Win32_OperatingSystem")
	
	 
	    For Each objItem in colOperatingSystems
	        OSVersion = left(objItem.Version,3)
	        ProductType = objItem.ProductType
	    Next
	
	
	    Select Case OSVersion
	        Case "6.1" 
	            OSName = "Windows 7"
	            ifFolderExist
	            copyTemplate
	        Case "5.2" 
	            OSName = "Windows 2003"
	            redFavorites()
	        Case "5.1" 
	            OSName = "Windows XP"
	            'ifFolderExist
	            redFavorites()
	        Case Else
	            OSName = "Hi! Grandpa! Windows ME or older"
	    End Select
	
	'Wscript.echo OSVersion & " " & ProductType
	    
	    
	
	    Set colItems = Nothing
	    Set objWMI = Nothing
End Sub

Function ifFolderExist()

Dim objFSO, strSourceFile, strTargetFolder, booOverWrite, strResult
Dim folders
Dim intFlag
Dim folder,envCar
strSourceFile = "\\S0300002\sysvol\D0300002.AOC\scripts\Msoffice\Jud11_Custom_Office_Ribbon_TabHome_Group.dotx"
strTargetFolder = "U:\My Documents\Word"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set envVar = WScript.CreateObject ("WScript.Shell")
Set folder = CreateObject("Scripting.FileSystemObject")
folders =Array("Word","Excel","Outlook","Powerpoint")
envVar.run "cmd /C cd c:\ & reg add ""HKCU\Software\Microsoft\Command Processor"" /v DisableUNCCheck /t REG_DWORD /d 1 /f "
'userPath = envVar.ExpandEnvironmentStrings("%UserProfile%")
userPath = envVar.ExpandEnvironmentStrings("u:\My Documents")
For Each f In folders 
	If folder.FolderExists(userPath & "\" & f) Then 
		'Script.Echo f & " " & " Exist"
		
	Else
	    'WScript.Echo f & " " & "doesnt Exist"
	    createDir f
	End If
Next

'If objFSO.FileExists( strTargetFile) Then
'strResult = "The file exists, overwritten"
'booOverWrite = vbTrue
'objFSO.CopyFile strSourceFile & "\*", strTargetFile & "\", booOverWrite
'Else
'strResult = "The file does not exist, created"
'booOverWrite = vbFalse
'objFSO.CopyFile strSourceFile, strTargetFile, booOverWrite
'End If
	    
End function




Sub deleteDir(folderName)
	Dim oShell
	Set oShell = WScript.CreateObject ("WScript.Shell")
	oShell.run "cmd /K  cd c:\ & RD %UserProfile%\" & foldername
	Set oShell = Nothing
End Sub

Sub redFavorites()
	Set objShell = WScript.CreateObject("WScript.Shell")
	RegLocate = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Favorites"
	objShell.RegWrite RegLocate,"U:\Favorites","REG_SZ"
	RegLocate = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Favorites"
	objShell.RegWrite RegLocate,"U:\Favorites","REG_SZ"
	WScript.Echo "done"
End Sub


Sub createDir(folderName)

	Dim oShell 
	Set oShell = WScript.CreateObject ("WScript.Shell")
	oShell.run "cmd /C  cd c:\ &  md ""u:\My Documents\""" &  folderName
	
	Set oShell = Nothing

End Sub

Sub copyTemplate

Dim objFSO, strSourceFile, strTargetFile, booOverWrite, strResult
strSourceFile = "\\S0300005\CITeS\Word Template\Jud11-Custom-Office-Ribbon-TabHome-Group.dotx"
strTargetFolder = "U:\My Documents\Word"
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists( strTargetFile) Then
strResult = "The file exists, overwritten"
booOverWrite = vbTrue
objFSO.CopyFile strSourceFile & "\*", strTargetFile & "\", booOverWrite
Else
strResult = "The file does not exist, created"
booOverWrite = vbFalse
objFSO.CopyFile strSourceFile, strTargetFolder & "\"
End If

'Wscript.Echo strResult
End Sub 
checkOs
