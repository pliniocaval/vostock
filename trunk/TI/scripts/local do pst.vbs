Set objOL = CreateObject("Outlook.Application")
Set objFolders = objOL.Session.Folders

For j = objFolders.Count To 1 Step -1
    Set objFolder = objFolders.Item(j)

    If (InStr(1, objFolder.Name, "Mailbox") = 0) And (InStr(1, objFolder.Name, "Public Folders") = 0) Then
	WScript.Echo objFolder.Name
	WScript.Echo GetPSTPath(objFolder.storeid)

End If
Next

Function GetPSTPath(input)
For i = 1 To Len(input) Step 2
	strSubString = Mid(input,i,2)
	If Not strSubString = "00" Then
		strPath = strPath & ChrW("&H" & strSubString)
	End If
Next

Select Case True
Case InStr(strPath,":\") > 0
	GetPSTPath = Mid(strPath,InStr(strPath,":\")-1)
Case InStr(strPath,"\\") > 0
	GetPSTPath = Mid(strPath,InStr(strPath,"\\"))
End Select
End Function