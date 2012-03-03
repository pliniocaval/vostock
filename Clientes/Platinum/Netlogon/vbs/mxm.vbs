'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 14/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

' Não parar em caso de erros
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

'Apagar o log se for maior que 10MB
If objFSO.FileExists(LOGUSER&"mxm.log") Then
set file = objFSO.GetFile(LOGUSER&"mxm.log")
  if file.Size >= 10485760 Then
    objFSO.DeleteFile(LOGUSER&"mxm.log")
  End If
End If

strTS = "MXM-CEMUSA"
'WScript.Echo computador
if UCASE(computador) = strTs Then
objnet.RemoveNetworkDrive "T:", true, true
objnet.MapNetworkDrive "T:", "\\csrv02\MX-Manager"
objFSO.DeleteFile Desktop&"\MXM*.lnk"
objFSO.DeleteFile Desktop&"\Microsoft Office O*.lnk"
objFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Iniciar o Navegador Internet Explorer.lnk"
objFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Microsoft Office O*.lnk"

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
objFSO.DeleteFile Desktop&"\MXM*.lnk"
Set MXMLnk = objShell.CreateShortcut(desktop & "\MXM.lnk")
MXMLnk.TargetPath = locmxm & "\MXMDV.RDP"
MXMLnk.Description = "Acesso remoto ao MXM"
MXMLnk.WorkingDirectory = locmxm
MXMLnk.WindowStyle = 1
MXMLnk.Save
wscript.quit

End if

