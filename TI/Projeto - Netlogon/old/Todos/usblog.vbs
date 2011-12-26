'Script do logon
'autoria Leonardo Vivas
'Versão 0.1
'criação 23/09/2010
'modificação 23/09/2010
' -----------------------------------------------------------------' 
Set objNetwork = CreateObject("Wscript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
strUserName = objNetwork.Username

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

' Não parar em caso de erros
On Error Resume Next

strLogFile = "\\csrv02\LOGS$\USB.log"

'inicio 
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE Targetinstance ISA 'Win32_PNPEntity' and TargetInstance.DeviceId like '%USBStor%'")
Do
Set objLatestEvent = colMonitoredEvents.NextEvent
Notifier(objLatestEvent.TargetInstance)
Loop

Sub Notifier(object)
Set objNet = CreateObject("Wscript.Network")

'You can change the function below to perform other actions
SendMailWithoutSSL _
"suporte@cemusadobrasil.com.br", _
"Dipositivo USB de Armazenamento detectado em " & objNet.Computername, _
"suporte@cemusadobrasil.com.br", _
"O usuario " & objNet.Username & " conectou um Dipositivo USB de Armazenamento em " & objNet.Computername & ".", _
"smtp.cemusadobrasil.com.br", _
25, _
"suporte@cemusadobrasil.com.br", _
"killer"

lRet = 2
Do While lRet = 2
   Msg = VbCrLf
   Msg = Msg & "Voce conectou um Dipositivo USB de Armazenamento." & chr(10) & VbCrLf
   Msg = Msg & "O uso deste tipo de dispositivo esta restrito aos cargos de Gerencia ou Superiores." & chr(10)& VbCrLf
   Msg = Msg & "O uso do dispositivo foi Registrado." & Chr(10)
   
lRet  =   MsgBox(msg,0,"Cemusa Informa")
Loop

arrTipos = split(arrTipos,";")
Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True, 0)
objLogFile.WriteLine  VBCRLF
objLogFile.WriteLine "==================================================="
objLogFile.WriteLine "O usuario " & objNet.Username & " conectou um Dipositivo USB de Armazenamento na estação " & objNet.Computername & " em "& now & "."
objLogFile.WriteLine "==================================================="
objLogFile.WriteLine  VBCRLF

End Sub




' CDOSYS official documentation: 
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wss/wss/_cdo_queue_top.asp
'
' by Vinicius Canto 
Sub SendMailWithoutSSL(strDestination, strTitle, strFrom, strMessage, strSMTP, intPort, strUsername, strPassword)
set oMessage = CreateObject("CDO.Message")
set oConf = CreateObject("CDO.Configuration")
Set oFields = oConf.Fields



oFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
oFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = intPort
oFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic: Auth with user and password sent with plain text
oFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUsername
oFields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPassword
oFields.Item("http://schemas.microsoft.com/cdo/configuration/Smtpusessl") = false
oFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1: Using local SMTP; 2: Using port; 3: Using Exchange
oFields.Update

oMessage.Fields.Item("urn:schemas:mailheader:to") = strDestination
oMessage.Fields.Item("urn:schemas:mailheader:from") = strFrom
oMessage.Fields.Item("urn:schemas:mailheader:sender") = strFrom 'reply-to
oMessage.Fields.Item("urn:schemas:mailheader:subject")= strTitle
oMessage.Fields.Item("urn:schemas:mailheader:x-mailer") = "Small Mail System -- by Leonardo Vivas "
oMessage.Fields.Update

oMessage.Configuration = oConf

oMessage.TextBody = strMessage
oMessage.Send
End Sub