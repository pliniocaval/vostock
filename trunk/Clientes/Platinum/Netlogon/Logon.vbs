'Script do logon
'autoria Leonardo Vivas
'Vers�o 2.0
'cria��o 03/06/2009
'modifica��o 03/03/2012
' -----------------------------------------------------------------' 

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'msgbox "N�o parar em caso de erros"
'On Error Resume Next

'MsgBox "Capturando Diretorio do Script"
DIR = oShell.CurrentDirectory

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
  
'msgbox "Remover drivers mapeados"
RemoveDrivesRede	

'msgbox "Criando pastas"
CriaPasta(TI)
CriaPasta(SUPORTE)
CriaPasta(HTA)
CriaPasta(PROGS)
CriaPasta(LOGS)
CriaPasta(USERLOGS)

'MsgBox "Tela de Logon"
TelaLogon

'msgbox "BGinfo"
BGINFO

'MsgBox "Executa Outros VBS"
