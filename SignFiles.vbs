Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Wscript.Shell")
objStartFolder = "C:\Signtool\In\"
objDestFolder = "C:\signtool\Out\"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files
i = 0
Do while i = 0
	On Error Resume Next
	For Each objFile in colFiles
	filesigned = False
    cml = "signtool.exe sign /f {CERTIFICATE.PFX} /p {PASSWORD} /fd sha256 /v /tr http://sha256timestamp.ws.symantec.com/sha256/timestamp " & objStartFolder & objFile.Name	
	'Do 
		intExitCode = objShell.Run(cml,0,True)
		If intExitCode = 0 Then
		WScript.Sleep 5000
		objFSO.DeleteFile objDestFolder & objFile.Name, True
		objFSO.MoveFile objStartFolder & objFile.Name, "C:\Signtool\Out\"
		filesigned = True
		Else 
		WScript.Sleep 20000
		End If 
	'Loop Until filesigned = True
	Next
WScript.Sleep 20000
Loop
Wscript.Quit
