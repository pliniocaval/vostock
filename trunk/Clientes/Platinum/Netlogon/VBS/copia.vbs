'Script Para Geração de Copia de Arquivos
'autoria Leonardo Vivas
'Versão 2.0
'criação 03/06/2009
'modificação 03/03/2012
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

'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = DIR & "\SYS\LOGON.INI"
  Set SYS = oFSO.OpenTextFile(varfile)
  SYSFILE =   SYS.ReadAll
  SYS.close
  execute SYSFILE

'msgbox "Carregando arquivo de Funções"
varfile = DIR & "\SYS\FNC.INI"
  Set FNC = oFSO.OpenTextFile(varfile)
  FNCFILE =   FNC.ReadAll
  FNC.close
  execute FNCFILE

'msgbox "Copia de arquivos"
If oFso.FileExists("C:\Windows\SysWOW64\RoboCopy.exe") Then
oFSO.DeleteFile WIN & "\system32\RoboCopy.exe"
Else
If oFso.FileExists("C:\Windows\System32\RoboCopy.exe") Then
'faça nada
Else
oFSO.CopyFile DIR & "\PROGS\RoboCopy.exe" , WIN & "\system32\RoboCopy.exe", OverwriteExisting
oFSO.CopyFile DIR&"\RoboCopy.exe" , PROGS & "\RoboCopy.exe", OverwriteExisting
End If
End If

'Wscript.Sleep 20000
'msgbox "Sync de Arquivos"
oShell.Run Robo & " " & DIR & "\HTA\ " & HTA & "\ " & RoboOPSYNC & USERLOGS & "\copyhta.log", 0, False
oShell.Run Robo & " " & DIR & "\IMG\ " & IMG & "\ " & RoboOPSYNC & USERLOGS & "\copyhta.log", 0, False
oShell.Run Robo & " " & DIR & "\PROGS\ " & PROGS & "\ " & RoboOPSYNC & USERLOGS & "\copyhta.log", 0, False
oShell.Run Robo & " " & DIR & "\SUPORTE\ " & SUPORTE & "\ " & RoboOPSYNC & USERLOGS & "\copysup.log", 0, False
'msgbox Fim
wscript.quit