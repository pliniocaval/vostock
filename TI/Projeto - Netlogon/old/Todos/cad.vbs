'Script do logon
'autoria Leonardo Vivas
'Vers�o 0.2
'cria��o 03/06/2009
'modifica��o 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' N�o parar em caso de erros
On Error Resume Next
strDom = objNet.UserDomain
strUser = objNet.UserName
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser &  ",user")
For Each objGroup In objUser.Groups
If objGroup.Name = "Diretoria" Then
 wscript.quit
End If
next

strLogFile = "c:\logs\"&objNet.UserName&"\cadastro.log"
set file = objFSO.GetFile(strLogFile)		
If DateDiff("d", file.DateLastModified, Now) > 360 Then 
objShell.Run "c:\suporte\Cadastro.hta",0 , True
else
 wscript.quit
end if
