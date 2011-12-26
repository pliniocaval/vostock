'Script do BKP OUTLOOK
'autoria Leonardo Vivas
'Versão 0.2
'criação 03/06/2009
'modificação 03/06/2009
' -----------------------------------------------------------------' 
Set objShell = CreateObject("WScript.Shell")
Set objnet = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
' Não parar em caso de erros
On Error Resume Next

'variaveis
UserName = objNet.Username
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\"
computador = objNet.ComputerName

if left(ucase(computador),4)="CSRV" then wscript.quit
if left(ucase(computador),3)="IMA" then wscript.quit
if left(ucase(computador),4)="VIRU" then wscript.quit
if left(ucase(computador),3)="TS2" then wscript.quit

outlook = "c:\outlook\"&UserName&"\"
bkpoutlook = "c:\outlook\BKP\"&UserName&"\"
logs = "c:\logs\"&UserName&"\"
locsuporte = "c:\suporte\"
suporte = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\softwares\"

'Apagar o log se for maior que 10MB
If objFSO.FileExists(logs&"outlook-bkp.log") Then
set file = objFSO.GetFile(logs&"outlook-bkp.log")
  if file.Size >= 10485760 Then
    objFSO.DeleteFile(logs&"outlook-bkp.log")
  End If
End If

'diretorios
objFSO.CreateFolder outlook
objFSO.CreateFolder "c:\outlook\BKP\"
objFSO.CreateFolder bkpoutlook
objFSO.CreateFolder "c:\logs\"
objFSO.CreateFolder logs

if left(ucase(computador),2)="TS" then 
'função de backup
robo = "c:\suporte\robocopy.exe "& Chr(34) & "c:\Documents and Settings\pvillaca\Configurações locais\Dados de aplicativos\Microsoft\Outlook\ "& Chr(34) & " c:\outlook\BKP\pvillaca\bkp01\ /TEE /S /E /COPY:DAT /R:100 /W:30 /LOG+:" & logs &"outlook-bkp1.log"
objFSO.CopyFile suporte&"RoboCopy.exe" , locsuporte&"RoboCopy.exe", True
set file = objFSO.GetFile(logs &"outlook-bkp1.log")		
If DateDiff("d", file.DateLastModified, Now) > 5 Then 
objShell.Run robo, 0, True
else
End if
End if

'função de backup
robo = "c:\suporte\robocopy.exe "& outlook &" "& bkpoutlook &" /TEE /S /E /COPY:DAT /R:10 /W:30 /LOG+:" & logs &"outlook-bkp.log"
objFSO.CopyFile suporte&"RoboCopy.exe" , locsuporte&"RoboCopy.exe", True
set file = objFSO.GetFile(logs &"outlook-bkp.log")		
If DateDiff("d", file.DateLastModified, Now) > 15 Then 
objShell.Run robo, 0, True
End If
wscript.quit