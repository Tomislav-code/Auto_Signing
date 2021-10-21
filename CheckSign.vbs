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
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "{MAILHOST@DOMAIN.COM}"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==

objMessage.Send
Set objMessage = Nothing
End Sub


mailto = "EMAIL_ADRESS@DOMAIN.COM"
mailfrom = "FROM_ADDRESS@DOMAIN.COM"


' verify signature
'WScript.Sleep 30000
cml = "signtool.exe verify /pa C:\signtool\Out\TestFileFromSignTool.exe" 
intExitCode = objShell.Run(cml,0,True)
	If intExitCode <> 0 Then
	SubjError = "Error: Checking signature Failed. "
	BodyText = "Verify signature failed, Please checkSignTool task" 
	sendmail SubjError, mailfrom, mailto, Bodytext	
	End If 
Wscript.Quit
