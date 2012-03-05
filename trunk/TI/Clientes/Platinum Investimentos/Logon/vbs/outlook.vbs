'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 14/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

' Não parar em caso de erros
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

if left(ucase(computador),3)="IMA" then wscript.quit
if left(ucase(computador),4)="CSRV" then wscript.quit
if left(ucase(computador),4)="VIRU" then wscript.quit
if left(ucase(computador),6)="CBSB04" then wscript.quit
if left(ucase(computador),7)="SQLSCPI" then wscript.quit
if left(ucase(computador),7)="CEMUSA-" then wscript.quit

'função de backup-profile-outlook
set file = objFSO.GetFile(LOGUSER &"\outlook-bkpprof.log")		
If DateDiff("d", file.DateLastModified, Now) > 30 Then
objFSO.DeleteFile LOGUSER &"\outlook-bkp1.log"
objFSO.DeleteFile LOGUSER &"\outlook-bkpprof.log"
objShell.Run "taskkill /F /IM outlook.exe", 0, True
objShell.Run bkpoutprof, 0, True
End if

'função de backup-outlook
set file = objFSO.GetFile(LOGUSER &"\outlook-bkp.log")		
If DateDiff("d", file.DateLastModified, Now) > 7 Then 
objFSO.DeleteFile LOGUSER &"\outlook-bkp.log"
objShell.Run "taskkill /F /IM outlook.exe", 0, True
objShell.Run bkpoutlook, 0, True
End If
wscript.quit