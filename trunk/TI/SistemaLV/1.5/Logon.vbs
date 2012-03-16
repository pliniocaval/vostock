'Script Para Logon
'autoria Leonardo Vivas
'Versão 2.0
'criação 03/06/2009
'modificação 03/03/2012
' -----------------------------------------------------------------' 

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'msgbox "Não parar em caso de erros"
'On Error Resume Next

'MsgBox "Capturando Diretorio do Script"
DIR = oFSO.GetParentFolderName(WScript.ScriptFullName)

'msgbox "Carregando variaveis"
  varfile = DIR & "\SYS\LOGON.INI"
  Set SYS = oFSO.OpenTextFile(varfile)
  SYSFILE =   SYS.ReadAll
  SYS.close
  execute SYSFILE

'msgbox "Carregando Arquivo de Funções"
varfile = DIR & "\SYS\FNC.INI"
  Set FNC = oFSO.OpenTextFile(varfile)
  FNCFILE =   FNC.ReadAll
  FNC.close
  execute FNCFILE
  
'msgbox "Carregando Arquivo de Parametrização"
varfile = DIR & "\SYS\EMP.INI"
  Set EMP = oFSO.OpenTextFile(varfile)
  EMPFILE =   EMP.ReadAll
  EMP.close
  execute EMPFILE

'msgbox "Criando pastas"
CriaPasta(TI)
CriaPasta(TIATU)
CriaPasta(HTA)
CriaPasta(IMG)
CriaPasta(PROGS)
CriaPasta(SUPORTE)
CriaPasta(SUPORTE & "\AUTOHELPDESK")
CriaPasta(SUPORTE & "\AUTOHELPDESK\INI")
CriaPasta(LOGS)
CriaPasta(USERLOGS)
CriaPasta(SRVLOG)

'MsgBox "Prepara de Arquivos Base"
DIRFILE = SUPORTE & "\AUTOHELPDESK\INI\DIRL.INI"
arrTipos = split(arrTipos,";")
Set DIRFILE = oFso.OpenTextFile(DIRFILE, 8, True, 0)
DIRFILE.WriteLine "DIRLOGON = " & Chr(34) & DIR & Chr(34)
CopiaArquivo DIR & "\SUPORTE\AUTOHELPDESK\INI\MAPS.INI" , SUPORTE & "\AUTOHELPDESK\INI\MAPS.INI"
CopiaArquivo DIR & "\IMG\Logo-Default.jpg" , IMG & "\Logo-Default.jpg"
CopiaArquivo DIR & "\HTA\Logon.hta",HTA & "\Logon.hta"

'MsgBox "Limpa Versão anterior do Script"
'ApagaRaiz(TIANT)
  
'msgbox "Remover drivers mapeados"
'RemoveDrivesRede

'MsgBox "Apaga logs Grandes demais"
ApagaArquivos2M(USERLOGS)

'msgbox "BGinfo"
BGINFO

'MsgBox "Tela de Logon"
TelaLogon(HTA & "\Logon.hta")

'MsgBox "Para Pelo Nome da Estação ou Servidor"
SAI

'MsgBox "Executa Outros VBS"
ExecutaVBS(DIR & "\VBS")

'MsgBox "Sincroniza arquivos de Log"
CopiaContPasta(USERLOGS)
