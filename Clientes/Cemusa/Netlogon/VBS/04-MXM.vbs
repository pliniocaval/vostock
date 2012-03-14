'Script Cemusa
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


strTS = "MXM-CEMUSA"
'WScript.Echo computador
if UCASE(COMP) = strTs Then
oNet.RemoveNetworkDrive "T:", true, true
oNet.MapNetworkDrive "T:", "\\csrv02\MX-Manager"
oFSO.DeleteFile DESK &"\MXM*.lnk"
oFSO.DeleteFile DESK &"\Microsoft Office O*.lnk"
oFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Iniciar o Navegador Internet Explorer.lnk"
oFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Microsoft Office O*.lnk"

lRet = 2
Do While lRet = 2
   Msg = VbCrLf
   Msg = Msg & "Voce esta no Terminal remoto do MXM." & chr(10) & VbCrLf
   Msg = Msg & "Todos os dias as 23:00 os arquivos salvos nesta maquina serão Apagados." & chr(10)& VbCrLf
   Msg = Msg & "Favor salvar os arquivos importantes na rede" & Chr(10)
   
lRet  =   MsgBox(msg,0,"Cemusa Informa")
Loop
wscript.quit
Else
objFSO.DeleteFile DESK &"\MXM*.lnk"
Set MXMLnk = objShell.CreateShortcut(DESK & "\MXM.lnk")
MXMLnk.TargetPath = MXM & "\MXMDV.RDP"
MXMLnk.Description = "Acesso remoto ao MXM"
MXMLnk.WorkingDirectory = MXM
MXMLnk.WindowStyle = 1
MXMLnk.Save
wscript.quit

End if

