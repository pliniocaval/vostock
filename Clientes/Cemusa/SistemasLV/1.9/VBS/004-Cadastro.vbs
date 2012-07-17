'Script do logon | Leonardo Vivas

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

if left(ucase(computador),4)="PVIL" then wscript.quitd
if left(ucase(computador),6)="CRJ010" then wscript.quitd

set file = oFSO.GetFile(CAD)		
If DateDiff("d", file.DateLastModified, Now) > 183 Then
CopiaArquivo  DIRS & "\HTA\Cadastro.hta",HTA & "\Cadastro.hta"
'wscript.sleep 5000
oShell.Run HTA & "\Cadastro.hta"
wscript.quit
else
wscript.quit
end if
