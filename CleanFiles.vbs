Dim fso, startFolder, OlderThanDate

Set fso = CreateObject("Scripting.FileSystemObject")
startFolder = "C:\signtool\Out" ' folder to start deleting (subfolders will also be cleaned)
OlderThanDate = DateAdd("d", -7, Date)  ' 7 days (adjust as necessary)

DeleteOldFiles startFolder, OlderThanDate

Function DeleteOldFiles(folderName, BeforeDate)
   Dim folder, file, fileCollection, folderCollection, subFolder

   Set folder = fso.GetFolder(folderName)
   Set fileCollection = folder.Files
   For Each file In fileCollection
      If file.DateLastModified < BeforeDate Then
         If UCase(fso.GetExtensionName(file.Name)) = "LOG" Then
         fso.DeleteFile(file.Path)
	 End If
      End If
   Next
End Function
