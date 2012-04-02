'Script Para Geração de Assinatura do Outlook | Leonardo Vivas
' -------------------------------------------------------------

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
oShell.Run Robo & " " & DIRS & "\IMG\ " & IMG & "\ " & RoboOPCOPY & USERLOGS & "\copyimg.log", 0, True
ApagaArquivosPastas(vAPPDATA &"\Microsoft\Signatures\") 
ApagaArquivosPastas(vAPPDATA &"\Microsoft\Assinaturas\")  

'msgbox "Carregando arquivo de Funções"
varfile = SYS & "\PARA\EMAIL.INI"
  Set EMAIL = oFSO.OpenTextFile(varfile)
  EMAILFILE =   EMAIL.ReadAll
  EMAIL.close
  execute EMAILFILE