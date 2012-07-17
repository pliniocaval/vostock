'Script De Logon | Leonardo Vivas

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'MsgBox "Não parar em caso de erros"
On Error Resume Next

'MsgBox "Capturando Diretorio do Script"
DIRS = oFSO.GetParentFolderName(WScript.ScriptFullName)

'MsgBox "Carregando Variaveis Remotas"
varfile = DIRS & "\SYS\DIRL.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'MsgBox "Verifica se é a Primeira vez"
If Not oFso.FolderExists(TI) Then
oShell.Run (DIRS & "\Prepara.vbs"),0 , True
End If

'MsgBox "Carregando Variaveis Locais"
varfile = SYS & "\VAR.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE
  
'MsgBox "Carregando Arquivo de Funções"
varfile = SYS & "\FNC.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'MsgBox "Carregando Arquivo de Parametrização"
varfile = SYS & "\PARA\EMP.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'MsgBox "Remover drivers mapeados"
RemoveDrivesRede

'MsgBox "Apaga logs Grandes demais"
ApagaArquivos2M(USERLOGS)

'MsgBox "Trava USB"
const HKEY_LOCAL_MACHINE = &H80000002
oShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start",4 ,"REG_DWORD"
oShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Modem\Start",4 ,"REG_DWORD"

'MsgBox "Tela de Logon"
TelaLogon(HTA & "\Logon.hta")

'MsgBox "ATUALIZA SCRIPT"
Set up = oFso.GetFile(LOGS & "\STARTUP.log")
If DateDiff("d", up.DateLastModified, Now) > 60 Then
oShell.Run (DIRS & "\Prepara.vbs"),0 , False
End If

'MsgBox "REALIZA INVENTARIO"
Set up = oFso.GetFile(USERLOGS & "\Inventario-" & COMP & ".log")
If DateDiff("d", up.DateLastModified, Now) > 60 Then
oShell.Run (DIRS & "\VBS\Inventario.vbs"),0 , True
CopiaContPasta USERLOGS,SRVLOG
End If

'MsgBox "Para Pelo Nome da Estação ou Servidor"
SAI

'MsgBox "BGinfo"
BGINFO

'MsgBox "Executa Outros VBS"
ExecutaVBS(DIRS & "\VBS")

'MsgBox "Criando pasta de Log Remota"
CriaPasta(SRVLOG)

'MsgBox "Sincroniza arquivos de Log"
CopiaContPasta USERLOGS,SRVLOG


'MsgBox "Limpa Versão anterior do Script"
ApagaRaiz(TIANT)

'MsgBox "Vereifica se executou corretamente"
If Not oFSO.DriveExists("H:") Then
oShell.Run (DIRS & "\SUPORTE\reparo.vbs"),0 ,True
End if
