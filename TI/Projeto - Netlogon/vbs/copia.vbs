'Script do logon
'autoria Leonardo Vivas
'Vers�o 1.8
'cria��o 03/06/2009
'modifica��o 14/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

Set objFSO = CreateObject("Scripting.FileSystemObject")
'msgbox "N�o parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

'msgbox "Copia de arquivos"
If objFso.FileExists("C:\Windows\SysWOW64\RoboCopy.exe") Then
objFSO.DeleteFile WIN & "\system32\RoboCopy.exe"
Else
If objFso.FileExists("C:\Windows\System32\RoboCopy.exe") Then
'fa�a nada
Else
objFSO.CopyFile instloc&"\RoboCopy.exe" , WIN & "\system32\RoboCopy.exe", OverwriteExisting
End If
End If

Wscript.Sleep 20000
'msgbox "Copia de arquivos - Hta"
objShell.Run CopyHta, 0, True
'msgbox "Copia de arquivos - MXM"
objShell.Run MXMCOPY, 0, True
'msgbox "Copia de arquivos - Install"
objShell.Run Inst, 0, True
'msgbox "Copia de arquivos - Remove"
objShell.Run Remove, 0, True
'msgbox "Copia de arquivos - Suporte"
objShell.Run CopySup, 0, True
'msgbox Fim
wscript.quit