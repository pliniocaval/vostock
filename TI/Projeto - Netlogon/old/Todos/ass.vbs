'Script do logon
'autoria Leonardo Vivas
'Versão 0.2
'criação 03/06/2009
'modificação 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Não parar em caso de erros
On Error Resume Next

'variaveis
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS"
suploc="c:\suporte\"
computador = objNet.ComputerName
vAPPDATA = objShell.ExpandEnvironmentStrings("%APPDATA%")

if left(ucase(computador),9)="CEMUSA003" then wscript.quit
if left(ucase(computador),4)="CSRV" then wscript.quit
if left(ucase(computador),4)="VIRU" then wscript.quit
if left(ucase(computador),3)="IMA" then wscript.quit
if left(ucase(computador),2)="TS" then wscript.quit

set folder = objFSO.getFolder (vAPPDATA &"\Microsoft\Signatures\")   
for each file in folder.files
if (dateDiff("d", file.datecreated, now) >3) then
File.delete
objfso.deletefolder vAPPDATA & "\Microsoft\Signatures\*.*",true
else
wscript.quit
end if
next

set folder = objFSO.getFolder (vAPPDATA &"\Microsoft\Assinaturas\")  
for each file in folder.files
if (dateDiff("d", file.datecreated, now) >3) then
File.delete
objfso.deletefolder vAPPDATA & "\Microsoft\Assinaturas\*.*",true
else
wscript.quit
end if
next

strDom = objNet.UserDomain
strUser = objNet.UserName
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser &  ",user")

Wscript.Sleep 5000

For Each objGroup In objUser.Groups

''''Scripts de email'''''
'assinaturas
 objShell.Run (scripts&"\todos\ass2-int.vbs"), 0, True
 objShell.Run (scripts&"\todos\ass2-ext.vbs"), 0, True
 Wscript.Sleep 5000
If objGroup.Name = "CEL" Then
 objShell.Run (scripts&"\todos\ass-int.vbs"), 0, True
 objShell.Run (scripts&"\todos\ass-ext.vbs"), 0, True
 End If
next
wscript.quit
