Set args = Wscript.Arguments

'Const strPath = 
'Var intDays = 
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim num

	
If (Wscript.Arguments.Item(2)=0) Then
Call SearchBefore (WScript.Arguments.Item(0),WScript.Arguments.Item(1))
Else
Call SearchBetween (Wscript.Arguments.Item(0), Wscript.Arguments.Item(1), Wscript.Arguments.Item(2))
End If


Sub SearchBefore(str,intDays)
    Dim objFolder, objSubFolder, objFile
    Set objFolder = objFSO.GetFolder(str)
    For Each objFile In objFolder.Files
 
        If objFile.DateCreated < (Now - intDays) Then
            objFile.Delete(True)
        End If
 
    Next
	
		
			
    For Each objSubFolder In objFolder.SubFolders
       Call SearchBefore(objSubFolder.Path,intDays)

		'objSubFolder.Files.Count = 0
	   If (Wscript.Arguments.Count>3) Then
			If objSubFolder.DateCreated < (Now - intDays) Then
            objFSO.DeleteFolder objSubFolder.Path
			
			End If
			'something=objSubFolder.Delete
       End If
    Next

   


End Sub

Sub SearchBetween(str,intDays, intDaysB)
    Dim objFolder, objSubFolder, objFile
    Set objFolder = objFSO.GetFolder(str)
    For Each objFile In objFolder.Files
 
        If (intDays > (DateDiff("d", Now, objFile.DateCreated)) & objFile.DateCreated > (Now - intDaysB)) Then
            objFile.Delete(True)
        End If
 
    Next
    For Each objSubFolder In objFolder.SubFolders
       Call SearchBetween(objSubFolder.Path,intDays,intDaysB)
 
    If (Wscript.Arguments.Count>3) Then
            If objSubFolder.DateCreated < (Now - intDays) & objSubFolder.DateCreated > (Now - intDaysB) Then
            objFSO.DeleteFolder objSubFolder.Path
			
			End If
			End If
 
    Next

End Sub

