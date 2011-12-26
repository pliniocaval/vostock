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

outlook = "c:\outlook\BKP\"&UserName&"\"
bkpoutlook = "\\csrv06\BKP$\"&UserName&"\outlook\BKP\"
logs = "\\csrv06\BKP$\"&UserName&"\outlook\"

'Apagar o log se for maior que 10MB
If objFSO.FileExists(logs&"outlook-bkp.log") Then
set file = objFSO.GetFile(logs&"outlook-bkp.log")
  if file.Size >= 10485760 Then
    objFSO.DeleteFile(logs&"outlook-bkp.log")
  End If
End If

'diretorios
objFSO.CreateFolder "\\csrv06\BKP$\"&UserName
objFSO.CreateFolder "\\csrv06\BKP$\"&UserName&"\outlook\"
objFSO.CreateFolder "\\csrv06\BKP$\"&UserName&"\outlook\BKP\"

'função de backup
robo = "c:\suporte\robocopy.exe "& outlook &" "& bkpoutlook &" /TEE /S /E /COPY:DAT /R:100 /W:30 /LOG+:" & logs &"outlook-bkp.log"

set file = objFSO.GetFile(logs &"outlook-bkp.log")		
If DateDiff("d", file.DateLastModified, Now) > 5 Then 
'wscript.echo File
'wscript.echo File.DateLastModified
objShell.Run robo, 0, True
End If
wscript.quit