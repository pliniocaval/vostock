'Script do logon
'autoria Leonardo Vivas
'Vers�o 0.1
'cria��o 03/06/2009
'modifica��o 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' N�o parar em caso de erros
On Error Resume Next

'variaveis
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\"

objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"

objShell.Run (scripts&"todos\const-logoff.vbs")

' Fim
WScript.Quit