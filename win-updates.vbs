'#
'# ServerPendingUpdates.vbs
'#
'# Usage: cscript ServerPendingUpdates.vbs {servername} {servername} {servername} {servername}
'#    If no {servername} specified then 'localhost' assumed
'#
'# To do: Error handling
'#
'#
'# Loop through the input parameters for each server
'#
'Dim i
'WScript.Echo strServer
'For i = 0 To WScript.Arguments.Count - 1

'  WScript.Echo strServer
'    CheckServerUpdateStatus GetArgValue(i,"localhost") 'strServer
'Next

getServers


'CheckServerUpdateStatus strServer 
WScript.Quit(0)


sub getServers()

Set objOU = GetObject("LDAP://OU=Server-group-low-impact,OU=Servers Group,OU=Computer Group,DC=D0300002,DC=AOC")
objOU.Filter = Array("Computer")
	For Each objComputer In objOU
		 
		CheckServerUpdateStatus(objComputer.CN)
	
	Next
	
	send_alert_email("c:\updateMonitor.log") 
		
End Sub  


Sub send_alert_email(strAttachment)
	Dim objEMail
	Set objEMail = CreateObject("CDO.Message")
	objEMail.From = "updates-Monitor@jud11.flcourts.org"
	objEMail.To = "lmendieta@jud11.flcourts.org"
	objEMail.Subject = "Windows Update  Monitor Alert"
	objEMail.TextBody = "Please see attached file for Information about Windows Updates on Low Impact Servers"
	
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

Function CheckServerUpdateStatus(strServer)
	
	
	
	Dim strFileName  : strFileName = "c:\updateMonitor.log"
	Dim strThisDay   : strThisDay=Date()
	Dim filesys       
	Dim filetxt 	  
	
	
	Set filesys =  CreateObject("Scripting.FilesystemObject")
	Set filetxt = filesys.OpenTextFile(strFileName, 8, True)
	WScript.Echo vbCRLF & "Connecting to " & strServer & " to check software update status..."
	Filetxt.WriteLine("******************************************************************************************")
	filetxt.Writeline(strThisDay & ":" & "Connecting to " & strServer & " to check software update status..." )
	
	
	On Error Resume next
    Dim blnRebootRequired    : blnRebootRequired     = False
    Dim blnRebootPending    : blnRebootPending     = False
    Dim objSession        : Set objSession    = CreateObject("Microsoft.Update.Session", strServer)
    Dim objUpdateSearcher     : Set objUpdateSearcher    = objSession.CreateUpdateSearcher
    Dim objSearchResult    : Set objSearchResult     = objUpdateSearcher.Search("IsInstalled = 0 and IsAssigned=1 and IsHidden=0 and Type='Software'")
    Dim i, objUpdate
    Dim intPendingInstalls    : intPendingInstalls     = 0

    For i = 0 To objSearchResult.Updates.Count - 1
        Set objUpdate = objSearchResult.Updates.Item(i)

        filetxt.WriteLine("-----"	&  " " & objUpdate.Title)
        If objUpdate.IsInstalled Then
            If objUpdate.RebootRequired Then
                blnRebootPending     = True
            End If
        Else
            intPendingInstalls    = intPendingInstalls + 1
            'If objUpdate.RebootRequired Then    '### This property is FALSE before installation and only set to TRUE after installation to indicate that this patch forced a reboot.
            If objUpdate.InstallationBehavior.RebootBehavior <> 0 Then
                '# http://msdn.microsoft.com/en-us/library/aa386064%28v=VS.85%29.aspx
                '# InstallationBehavior.RebootBehavior = 0    Never reboot
                '# InstallationBehavior.RebootBehavior = 1    Must reboot
                '# InstallationBehavior.RebootBehavior = 2    Can request reboot
                blnRebootRequired     = True
            End If

        End If
    Next
    

	If intPendingInstalls > 0 Then 
	WScript.Echo  "-----"	& " " & strServer & " has " & intPendingInstalls & " updates pending installation"
    filetxt.WriteLine("-----"	& " " &	strServer & " has " & intPendingInstalls & " updates pending installation")

	    If blnRebootRequired Then
	        WScript.Echo "-----"	& " " & strServer & " WILL need to be rebooted to complete the installation of these updates."
	        filetxt.WriteLine(	"-----"	& " " &	strServer & " WILL need to be rebooted to complete the installation of these updates.")
	        
	    Else
	        WScript.Echo  "-----"	& strServer & " WILL NOT require a reboot to install these updates."
	        filetxt.WriteLine( "-----"	&		strServer & " WILL NOT require a reboot to install these updates.")
	    End If
	
	
	    If blnRebootPending Then
	        WScript.Echo strServer & " is waiting for a REBOOT to complete a previous installation."
	        filetxt.WriteLine(		strServer & " is waiting for a REBOOT to complete a previous installation.")
	    End If 
	Else 
	
	WScript.Echo "Updates fully Installed"
	filetxt.WriteLine("Updates fully Installed")	
		
	End If




    
    
    filetxt.Close()
    
    
End Function



