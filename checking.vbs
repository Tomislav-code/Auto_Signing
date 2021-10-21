Set objShell = CreateObject("Wscript.Shell")
Set objFSO=CreateObject("Scripting.FileSystemObject")

Sub sendmail(subj,mfrom,mto,body) 
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = subj
objMessage.From = mfrom
objMessage.To = mto
objMessage.TextBody = body

'==This section provides the configuration information for the remote SMTP server.
'==Normally you will only change the server name or IP.

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "{SMTP SERVER ADDRESS}.com"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==

objMessage.Send
Set objMessage = Nothing
End Sub


mailto = "EMAIL_ADDRESS@DOMAIN.COM"
mailfrom = "FROM_ADDRESS@DOMAIN.COM"
strFileName = "TestFileFromSignTool.exe"
On Error Resume Next
'intExitCode = objFSO.CopyFile(strFileName, "C:\Signtool\In\", False)
objFSO.CopyFile strFileName, "C:\Signtool\In\", False
	
	If Err.Number <> 0 Then
	WScript.Sleep 5000
	'Email notification error
	SubjError = "Error: Checking SignTool Failed. "
	BodyText = "TestFileFromSignTool.exe is stuck, Please restart SignTool task" 
	sendmail SubjError, mailfrom, mailto, Bodytext	
	End If 

Wscript.Quit
