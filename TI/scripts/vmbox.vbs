
' ----- ExeScript Options Begin -----
' ScriptType: window,invoker
' DestDirectory: temp
' Icon: default
' ----- ExeScript Options End -----
'Script para VMBOX
'autoria Leonardo Vivas
'Versão 0.3
'criação 03/06/2009
'modificação 12/07/2010
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Não parar em caso de erros
On Error Resume Next

pedro = "c:\progra~1\Oracle\VirtualBox\VBoxManage.exe startvm --type gui 5881262c-e4a0-4ee2-a4d7-0a3f5972c8c5"
FTP = "c:\progra~1\Oracle\VirtualBox\VBoxManage.exe startvm --type gui 9d0e88fa-b9bb-4211-b0e5-924fd6d2cb62"

objShell.Run pedro
Wscript.Sleep 120000
objShell.Run FTP