'Script Para Gera��o de Assinatura para o Outlook
'autoria Leonardo Vivas
'Vers�o 2.0
'cria��o 03/06/2009
'modifica��o 03/03/2012
' -----------------------------------------------------------------' 

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oSysInfo = CreateObject("ADSystemInfo")
'Captura e volta 1 nivel de diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")
For i = 0 to Ubound(arrPath) - 1
    DIR = DIR & arrPath(i) & "\"
Next 

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

ApagaArquivosPastas(vAPPDATA &"\Microsoft\Signatures\") 
ApagaArquivosPastas(vAPPDATA &"\Microsoft\Assinaturas\")  

'msgbox "Carregando arquivo de Fun��es"
varfile = DIR & "SUPORTE\AUTOHELPDESK\INI\EMAIL.INI"
  Set EMAIL = oFSO.OpenTextFile(varfile)
  EMAILFILE =   EMAIL.ReadAll
  EMAIL.close
  execute EMAILFILE