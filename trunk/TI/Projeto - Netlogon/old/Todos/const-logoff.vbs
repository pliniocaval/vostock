' Author Leonardo Vivas
' Versão 0.5 - Maio 2009
' -----------------------------------------------------------------' 
'Pega informaçoes do Usuario
Set objNet = CreateObject("WScript.Network")
Set objNetwork = CreateObject("Wscript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS"

' Não para em Caso de Erros
On Error Resume Next

'bkp local outlook
objShell.Run (scripts&"\todos\outlook.vbs")

'Registra Desligamento
objShell.Run (scripts&"\todos\log2.vbs"), 0, True

strDom = objNet.UserDomain
strUser = objNet.UserName
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser &  ",user")

For Each objGroup In objUser.Groups

next
Set objFSO = Nothing
Set objNetwork = Nothing
Set objShell = Nothing

WScript.Quit