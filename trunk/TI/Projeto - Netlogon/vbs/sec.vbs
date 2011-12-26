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

'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\Logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes = f.ReadAll
  f.close
  execute constantes
  
'msgbox "Alterando ACL"
objFSO.DeleteFile LOGUSER & "\xcalcs*.log"
objShell.Run (scripts&"\vbs\XCACLS.vbs " & ti & " /g Everyone:f /f /t /q /l " & LOGUSER & "\xcalcs-ti.log"),0 , False
objShell.Run (scripts&"\vbs\XCACLS.vbs " & outlookuser & " /g " & Domain & "\" & user & ":f /f /t /q /l " & LOGUSER & "\xcalcs-outlook.log"),0 , True
objShell.Run (scripts&"\vbs\XCACLS.vbs " & outlookbkpuser & " /g " & Domain & "\" & user & ":f /f /t /q /l " & LOGUSER & "\xcalcs-outlookbkp.log"),0 , True
objShell.Run (scripts&"\vbs\XCACLS.vbs " & outlookuser & " /e /g cemusa\BKP:f /f /t /q /l " & LOGUSER & "\xcalcs-BKP.log"),0 , True
objShell.Run (scripts&"\vbs\XCACLS.vbs " & outlookbkp & " /e /g cemusa\BKP:f /f /t /q /l " & LOGUSER & "\xcalcs-BKP.log"),0 , True
objShell.Run (scripts&"\vbs\XCACLS.vbs " & outlookbkpuser & " /e /g cemusa\BKP:f /f /t /q /l " & LOGUSER & "\xcalcs-BKP.log"),0 , True
objShell.Run (scripts&"\vbs\XCACLS.vbs " & outlookrede & " /e /g cemusa\BKP:f /f /t /q /l " & LOGUSER & "\xcalcs-BKP.log"),0 , True