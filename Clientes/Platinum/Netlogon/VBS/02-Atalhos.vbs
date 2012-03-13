'Script Para Geração de Atalhos no desktop
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

'MsgBox "Atalho para Auto Help Desk "  
Set ReparoLnk = oShell.CreateShortcut(DESK & "\Auto Help Desk.lnk")
ReparoLnk.TargetPath = suploc&"\HelpDesk.hta"
ReparoLnk.Description = "Auto Help Desk"
ReparoLnk.WorkingDirectory = HTA
ReparoLnk.WindowStyle = 1
ReparoLnk.IconLocation = htaloc &"\img\logo.ico"
ReparoLnk.Save

'msgbox "Atalho no Desktop para a Rede."
Set DepLnk = oShell.CreateShortcut(DESK & "\Departamentos " & DOMI & ".lnk")
DepLnk.TargetPath = "M:\"
DepLnk.Description = "Atalho para " & DOMI
DepLnk.WorkingDirectory = "M:\"
DepLnk.WindowStyle = 1
DepLnk.Save
