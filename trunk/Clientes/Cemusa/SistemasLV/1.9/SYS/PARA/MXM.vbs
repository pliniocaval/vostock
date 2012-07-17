'Script Para Publicação do MXM via TS | Leonardo Vivas
' -------------------------------------------------------------

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Captura e volta 1 nivel do diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")
For i = 0 to Ubound(arrPath) - 2
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

  
strTS = "MXM-CEMUSA"
'WScript.Echo computador
if UCASE(COMP) = strTs Then
oNet.RemoveNetworkDrive "T:", true, true
oNet.MapNetworkDrive "T:", "\\csrv02\MX-Manager"
oFSO.DeleteFile DESK &"\MXM*.lnk"
oFSO.DeleteFile DESK &"\Microsoft Office O*.lnk"
oFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Iniciar o Navegador Internet Explorer.lnk"
oFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Microsoft Office O*.lnk"
oShell.Run Chr(34) & PROG & "\CCleaner\ccleaner.exe" & Chr(34) & " /AUTO",0 ,False
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
oFSO.DeleteFile DESK &"\MXM*.lnk"
oShell.Run Robo & " " & SRVMXM & " " & MXM & "\ " & RoboOPCOPY & USERLOGS & "\copyMXM.log", 0, True
Set MXMLnk = oShell.CreateShortcut(DESK & "\MXM.lnk")
MXMLnk.TargetPath = MXM & "\MXMDV.RDP"
MXMLnk.Description = "Acesso remoto ao MXM"
MXMLnk.WorkingDirectory = MXM
MXMLnk.WindowStyle = 1
MXMLnk.Save
wscript.quit

End if

