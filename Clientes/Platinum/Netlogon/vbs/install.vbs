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

'msgbox "N�o parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\Logon.ini"
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes
  
'vers�o do SO
set objwmiservice = GetObject("winmgmts:\\")
Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colComputer
PC_Type = objComputer.SystemType
next
'MsgBox PC_Type

'msgbox "Instala Arquivos"
'msgbox "Carregando variaveis"
progfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\inst\progs.ini"
  Set f = objFSO.OpenTextFile(progfile)
  constantes =   f.ReadAll
  f.close
  execute constantes
  
If PC_Type = "x86-based PC" Then 
If Not objFso.FileExists(LOGLOC & "\vcredist_x86.log") Then objShell.Run (psexec & " " & pstoolsvar & " " & vcredistX86),0 ,True
objFSO.CreateTextFile LOGLOC & "\vcredist_x86.log",true
'If Not objFso.FileExists(LOGLOC & "\eset.log") Then objShell.Run (psexec & " " & pstoolsvar & " " & esetX86),0 ,True
'objFSO.CreateTextFile LOGLOC & "\eset.log",true
End If
'X64-based PC
If PC_Type = "x64-based PC" Then
If Not objFso.FileExists(LOGLOC & "\vcredist_x86.log") Then objShell.Run (psexec & "  " & pstoolsvar & "  " & vcredistX64),0 ,True
objFSO.CreateTextFile LOGLOC & "\vcredist_x86.log",true
'If Not objFso.FileExists(LOGLOC & "\eset.log") Then objShell.Run (psexec & " " & pstoolsvar & " " & esetX64),0 ,True
'objFSO.CreateTextFile LOGLOC & "\eset.log",true
End If

'msgbox Fim
wscript.quit