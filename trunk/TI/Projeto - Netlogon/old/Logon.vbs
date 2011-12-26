'Script do logon
'autoria Leonardo Vivas
'Versão 1.0
'criação 03/06/2009
'modificação 08/02/2011
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const OverwriteExisting = True

' Não parar em caso de erros
On Error Resume Next

'variaveis
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\"
computador = objNet.ComputerName
BgInfo = "c:\suporte\inst\bginfo.exe c:\suporte\cemusa.bgi /timer:0 /nolicprompt"

'Registro, Impressora
objShell.Run (scripts&"todos\reg.vbs")
'objShell.Run (scripts&"todos\outlook.vbs"),0 , False
'objShell.Run (scripts&"todos\print.vbs")

'Block UBS
const HKEY_LOCAL_MACHINE = &H80000002
objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start",4 ,"REG_DWORD"
objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Modem\Start",4 ,"REG_DWORD"

'Tela de Logon
objFSO.CopyFile scripts&"software\logon.hta" , "c:\suporte\logon.hta", OverwriteExisting
objShell.Run ("c:\suporte\logon.hta"),1 ,False
'rotina de copia
objShell.Run (scripts&"todos\copy.vbs"),0 ,True
'update de Progamas
objShell.Run (scripts&"todos\install.vbs"),1 ,True
'criticas de saida 01
if left(ucase(computador),6)="CSRV04" then wscript.quit
'Registro de Logon
objShell.Run (scripts&"todos\log.vbs"),0 , False
'Critica de saida
if left(ucase(computador),2)="TS" then wscript.quit
if left(ucase(computador),3)="IMA" then wscript.quit
if left(ucase(computador),4)="CSRV" then wscript.quit
if left(ucase(computador),4)="VIRU" then wscript.quit
if left(ucase(computador),6)="CBSB04" then wscript.quit
if left(ucase(computador),7)="SQLSCP1" then wscript.quit
if left(ucase(computador),7)="CEMUSA-" then wscript.quit
'cadastramento de usuario 
objShell.Run (scripts&"todos\cad.vbs"),0 , True
'assinatura de email
objShell.Run (scripts&"todos\ass.vbs"),0 , True
'BGinfo
objShell.Run BgInfo,0 , False
