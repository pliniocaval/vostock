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

' Não parar em caso de erros
On Error Resume Next

'Carregando variaveis
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

'Apagar o log se for maior que 10MB
If objFSO.FileExists(LOGUSER & "\uninst.log") Then
set file = objFSO.GetFile(LOGUSER & "\uninst.log")
  if file.Size >= 10485760 Then
    objFSO.DeleteFile(LOGUSER & "\uninst.log")
  End If
End If

'versão do SO
set objwmiservice = GetObject("winmgmts:\\")
Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colComputer
PC_Type = objComputer.SystemType
next
'MsgBox PC_Type

'msgbox "Instala Arquivos"
'msgbox "Carregando variaveis"
progfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\uninst\progs.ini"
  Set f = objFSO.OpenTextFile(progfile)
  constantes =   f.ReadAll
  f.close
  execute constantes
 
If PC_Type = "x86-based PC" Then 
'MsgBox psexec & "  " & pstoolsvar & "  " & gtalkx86
If Not objFso.FileExists(LOGLOC & "\gtalkx86.log") Then objShell.Run (psexec & " " & pstoolsvar & " " & gtalkx86),0 ,True
End If
'X64-based PC
If PC_Type = "x64-based PC" Then
'MsgBox psexec & "  " & pstoolsvar & "  " & gtalkx64
If Not objFso.FileExists(LOGLOC & "\gtalkx64.log") Then objShell.Run (psexec & "  " & pstoolsvar & "  " & gtalkx64),0 ,True
End If

'apaga diretorios
Set objFolder = objFSO.GetFolder("C:\logs")
objFolder.Delete True
Set objFolder = objFSO.GetFolder("C:\suporte")
objFolder.Delete True
Set objFolder = objFSO.GetFolder("C:\mxm")
objFolder.Delete True
wscript.quit