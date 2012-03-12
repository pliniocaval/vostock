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


'MsgBox "Sync Pastas"
oFSO.CopyFolder DIR & "\HTA\*.*", HTA & "\" , OverwriteExisting
oFSO.CopyFolder DIR & "\IMG\*.*", IMG & "\" , OverwriteExisting
oFSO.CopyFolder DIR & "\PROGS\*.*", PROGS & "\" , OverwriteExisting
oFSO.CopyFolder DIR & "\SUPORTE\*.*", SUPORTE & "\" , OverwriteExisting

'msgbox "Sync de Arquivos"
oFSO.CopyFile DIR & "\HTA\*.*", HTA & "\" , OverwriteExisting
oFSO.CopyFile DIR & "\IMG\*.*", IMG & "\" , OverwriteExisting
oFSO.CopyFile DIR & "\PROGS\*.*", PROGS & "\" , OverwriteExisting
oFSO.CopyFile DIR & "\SUPORTE\*.*", SUPORTE & "\" , OverwriteExisting

'msgbox Fim
wscript.quit