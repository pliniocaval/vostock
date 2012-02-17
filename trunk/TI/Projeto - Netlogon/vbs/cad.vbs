'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 21/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

' Não parar em caso de erros
On Error Resume Next

'Carregando variaveis
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

if left(ucase(computador),4)="PVIL" then wscript.quit

set file = objFSO.GetFile(cadLogFile)		
If DateDiff("d", file.DateLastModified, Now) > 183 Then 
objShell.Run htaloc&"\Cadastro.hta",0 , True
wscript.quit
else
wscript.quit
end if
