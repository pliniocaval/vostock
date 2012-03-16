'Script Para Gera��o de Copia de Arquivos
'autoria Leonardo Vivas
'Vers�o 2.0
'cria��o 03/06/2009
'modifica��o 03/03/2012
' -----------------------------------------------------------------' 

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Captura e volta 1 nivel do diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")

For i = 0 to Ubound(arrPath) - 1
    DIR = DIR & arrPath(i) & "\"
Next 

oShell.CurrentDirectory = DIR

'msgbox "N�o parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = DIR & "\SYS\LOGON.INI"
  Set SYS = oFSO.OpenTextFile(varfile)
  SYSFILE =   SYS.ReadAll
  SYS.close
  execute SYSFILE

'msgbox "Carregando arquivo de Fun��es"
varfile = DIR & "\SYS\FNC.INI"
  Set FNC = oFSO.OpenTextFile(varfile)
  FNCFILE =   FNC.ReadAll
  FNC.close
  execute FNCFILE
  
'msgbox "Carregando arquivo de Parametriza��o"
varfile = DIR & "\SYS\EMP.INI"
  Set EMP = oFSO.OpenTextFile(varfile)
  EMPFILE =   EMP.ReadAll
  EMP.close
  execute EMPFILE

'msgbox "Copia de arquivos"
If oFso.FileExists("C:\Windows\SysWOW64\RoboCopy.exe") Then
oFSO.DeleteFile WIN & "\system32\RoboCopy.exe"
Copia DIR & "\PROGS\RoboCopy.exe" , PROGS & "\RoboCopy.exe"
Else
If oFso.FileExists("C:\Windows\System32\RoboCopy.exe") Then
Copia DIR & "\PROGS\RoboCopy.exe" , PROGS & "\RoboCopy.exe"
Else
Copia DIR & "\PROGS\RoboCopy.exe" , PROGS & "\RoboCopy.exe"
Robo = PROGS & "\RoboCopy.exe"
End If
End If


'msgbox "Sync de Arquivos"
'SyncFiles
oShell.Run Robo & " " & DIR & "\HTA\ " & HTA & "\ " & RoboOPSYNC & USERLOGS & "\copyhta.log", 0, False
oShell.Run Robo & " " & DIR & "\IMG\ " & IMG & "\ " & RoboOPSYNC & USERLOGS & "\copyimg.log", 0, False
oShell.Run Robo & " " & DIR & "\PROGS\ " & PROGS & "\ " & RoboOPSYNC & USERLOGS & "\copyprg.log", 0, False
oShell.Run Robo & " " & DIR & "\SUPORTE\ " & SUPORTE & "\ " & RoboOPSYNC & USERLOGS & "\copysup.log", 0, False
oShell.Run Robo & " " & DIR & "\ " & TIATU & "\ " & "/XD VBS LOGS " & RoboOPSYNC & USERLOGS & "\copyTI.log", 0, False
'msgbox Fim
wscript.quit