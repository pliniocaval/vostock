'Script De Logon | Leonardo Vivas

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'msgbox "N�o parar em caso de erros"
'On Error Resume Next

'MsgBox "Capturando Diretorio do Script"
DIRS = oFSO.GetParentFolderName(WScript.ScriptFullName)

'msgbox "Carregando Variaveis Remotas"
varfile = DIRS & "\SYS\DIRL.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

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

'MsgBox "Limpa Vers�o anterior do Script"
'ApagaRaiz(TIANT)

  
'msgbox "Remover drivers mapeados"
'RemoveDrivesRede

'MsgBox "Apaga logs Grandes demais"
ApagaArquivos2M(USERLOGS)

'msgbox "BGinfo"
BGINFO

'MsgBox "Tela de Logon"
TelaLogon(HTA & "\Logon.hta")

'MsgBox "Para Pelo Nome da Esta��o ou Servidor"
SAI

'MsgBox "Executa Outros VBS"
ExecutaVBS(DIRS & "\VBS")

'msgbox "Criando pasta de Log Remota"
CriaPasta(SRVLOG)

'MsgBox "Sincroniza arquivos de Log"
CopiaContPasta(USERLOGS)