'Script do logon
'autoria Leonardo Vivas
'Versão 0.2
'criação 03/06/2009
'modificação 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
' Não parar em caso de erros
On Error Resume Next

'variaveis
UserName = objNet.Username
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\"

suporte = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\softwares\"
locsuporte = "c:\suporte\"
logs = "C:\logs\"

'Apagar o log se for maior que 10MB
If objFSO.FileExists(logs&"copy.log") Then
set file = objFSO.GetFile(logs&"copy.log")
  if file.Size >= 10485760 Then
    objFSO.DeleteFile(logs&"copy.log")
  End If
End If

'diretorios
objFSO.CreateFolder locsuporte
objFSO.CreateFolder logs

'função de backup
robo = "c:\suporte\robocopy.exe "& suporte &" "& locsuporte &" /MIR /TEE /LOG+:" & logs&"copy.log"

objFSO.CopyFile suporte&"RoboCopy.exe" , locsuporte&"RoboCopy.exe", OverwriteExisting

'set file = objFSO.GetFile(logs &"copy.log")		
'If DateDiff("d", file.DateLastModified, Now) > 5 Then 
'wscript.echo File
'wscript.echo File.DateLastModified
objShell.Run robo, 0, True
'End If
wscript.quit