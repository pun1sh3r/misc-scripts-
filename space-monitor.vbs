' List Available Disk Space

get_servers_from_file

Sub send_alert_email(strAttachment)
  Dim objEMail
	Set objEMail = CreateObject("CDO.Message")
	objEMail.From = "diskmon@jud11.flcourts.org"
	objEMail.To = "lmendieta@jud11.flcourts.org"
	objEMail.Subject = "Disk Monitor Alert"
	objEMail.TextBody = "Please see attached file for Server Disk Info"
	
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
	        "s0300088.D0300002.AOC" 
	objEmail.Configuration.Fields.Item _
	    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEMail.AddAttachment(strAttachment)
	objEMail.send

End Sub 

Sub get_servers_from_file()
	Dim filesys, filetxt
	'Dim dataArray
	Const HARD_DISK = 3
	Dim	strdata
	Dim strThisDay
	Dim strname
	Dim strFileName
	Dim dblFreeSpace
	Dim intCounter 
	
	Const ForReading = 1
	intCounter = 0
	strThisDay=Right(Year(Date),2) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2)
	strFileName = "c:\" & strThisDay & "StorageReport.txt"
	Set filesys = CreateObject("Scripting.FilesystemObject")
	Set filetxt = filesys.OpenTextFile( strFileName,2,True)
	Set objFSO = CreateObject("Scripting.FilesystemObject")
	Set objTextFile = objFSO.OpenTextFile("U:\servers.csv", ForReading)
	
	filetxt.WriteLine("*******************************************************************************")
	filetxt.WriteLine("*                        SERVER  DISK MONITOR                                 *")
	filetxt.WriteLine("*                                                                             *")
	filetxt.WriteLine("*               designed to monitor and report disk usage                     *")
	filetxt.WriteLine("*                     developed by Luis Mendieta                              *")
	filetxt.WriteLine("*******************************************************************************")
		Do until objTextFile.AtEndOfStream
			strdata = objTextFile.ReadLine
			arrayList = Split(strdata , ",")
			strname = arrayList(0)
			strComputer = strname
			On Error Resume next
			Set objWMIService = GetObject("winmgmts:" _
			    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			
			Set colDisks = objWMIService.ExecQuery _
			    ("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")
			For Each objDisk in colDisks
			
			dblFreeSpace = FormatNumber(objDisk.FreeSpace / 1048576000, 2)	     
				If dblFreeSpace < 1.0 Then
					intCounter = intCounter + 1 
					filetxt.WriteLine("-----------------------------------------")
					filetxt.WriteLine(strcomputer)
				    filetxt.WriteLine ("DeviceID: "& vbTab &  objDisk.DeviceID)       
				    filetxt.WriteLine ("Free Disk Space: "& vbTab & dblFreeSpace)
				    
				Else 
					
				End If 
			Next
		Loop
		
	filetxt.Close
	objTextFile.Close()
	If intCounter >= 2 Then
		send_alert_email(strFileName)
	Else 
		Exit Sub 
	End If
	
End Sub 
