'Script Para Sync de Arquivos | Leonardo Vivas
' ----------------------------------------------

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Captura e volta 1 nivel do diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")
For i = 0 to Ubound(arrPath) - 1
    DIRS = DIRS & arrPath(i) & "\"
Next 
oShell.CurrentDirectory = DIRS

'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando Variaveis Remotas"
DIRLfile = DIRS & "\SYS\DIRL.INI"
  Set DIRL = oFSO.OpenTextFile(DIRLfile)
  DIRLFILE =   DIRL.ReadAll
  DIRL.close
  execute DIRLFILE

'msgbox "Carregando Variaveis Locais"

varfile = SYS & "\VAR.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'msgbox "Carregando Arquivo de Funções"
varfile = SYS & "\FNC.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'msgbox "Carregando Arquivo de Parametrização"
varfile = SYS & "\PARA\EMP.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'msgbox "Copia de arquivos"
If oFso.FileExists("C:\Windows\SysWOW64\RoboCopy.exe") Then
Else
If oFso.FileExists("C:\Windows\System32\RoboCopy.exe") Then
Else
If oFso.FileExists(PROGS & "\RoboCopy.exe") Then
Else
CopiaArquivo DIRS & "\PROGS\RoboCopy.exe",PROGS & "\RoboCopy.exe"
Robo = PROGS & "\RoboCopy.exe"
End If
End If
End If

'msgbox "Criando pastas"
CriaPasta(TI)
CriaPasta(TIATU)
CriaPasta(HTA)
CriaPasta(IMG)
CriaPasta(PROGS)
CriaPasta(LOGS)
CriaPasta(SYS)
CriaPasta(SUPORTE)
CriaPasta(USERLOGS)

'msgbox "Sync de Arquivos"
oShell.Run Robo & " " & DIRS & "\HTA\ " & HTA & "\ " & RoboOPSYNC & USERLOGS & "\copyhta.log", 0, False
oShell.Run Robo & " " & DIRS & "\IMG\ " & IMG & "\ " & RoboOPSYNC & USERLOGS & "\copyimg.log", 0, False
oShell.Run Robo & " " & DIRS & "\PROGS\ " & PROGS & "\ " & RoboOPSYNC & USERLOGS & "\copyprg.log", 0, False
oShell.Run Robo & " " & DIRS & "\SYS\ " & SYS & "\ " & RoboOPSYNC & USERLOGS & "\copysys.log", 0, False
oShell.Run Robo & " " & DIRS & "\SUPORTE\ " & SUPORTE & "\ " & RoboOPCOPY & USERLOGS & "\copysup.log", 0, False
'msgbox Fim
wscript.quit

Function Loadfile(File)
  varfile = File
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE
End Function
