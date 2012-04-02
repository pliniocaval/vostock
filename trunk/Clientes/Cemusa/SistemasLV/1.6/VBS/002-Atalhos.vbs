'Script Para Gera��o de Atalhos | Leonardo Vivas
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

'msgbox "N�o parar em caso de erros"
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

'msgbox "Carregando Arquivo de Fun��es"
varfile = SYS & "\FNC.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'msgbox "Carregando Arquivo de Parametriza��o"
varfile = SYS & "\EMP.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'Limpa vers�o anterior
oFSO.DeleteFile DESK & "\Auto Help Desk.lnk"
oFSO.DeleteFile DESK & "\Departamentos " & DOMI & ".lnk"

'MsgBox "Atalho no Desktop para a Rede."
Set DepLnk = oShell.CreateShortcut(DESK & "\Departamentos " & DOMI & ".lnk")
DepLnk.TargetPath = "Y:\"
DepLnk.Description = "Atalho para " & DOMI
DepLnk.WorkingDirectory = "Y:\"
DepLnk.WindowStyle = 1
DepLnk.Save

'MsgBox "Atalho para Auto Help Desk "  
'Set ReparoLnk = oShell.CreateShortcut(DESK & "\Auto Help Desk.lnk")
'ReparoLnk.TargetPath = SUPORTE & "\HelpDesk.hta"
'ReparoLnk.Description = "Auto Help Desk"
'ReparoLnk.WorkingDirectory = HTA
'ReparoLnk.WindowStyle = 1
'ReparoLnk.IconLocation = IMG &"\logo.ico"
'ReparoLnk.Save