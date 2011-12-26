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

Set objFSO = CreateObject("Scripting.FileSystemObject")
'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

'msgbox "Copia de arquivos"
objFSO.CopyFile scripts&"\suporte\RoboCopy.exe" , robocopy, OverwriteExisting
'msgbox "Copia de arquivos - Suporte"
objShell.Run CopySup, 0, True
'msgbox "Copia de arquivos - Hta"
objShell.Run CopyHta, 0, True
'msgbox "Copia de arquivos - MXM"
objShell.Run MXMCOPY, 0, True
'msgbox "Copia de arquivos - Install"
objShell.Run Inst, 0, True
'msgbox "Copia de arquivos - Remove"
objShell.Run Remove, 0, True

'msgbox Fim
wscript.quit