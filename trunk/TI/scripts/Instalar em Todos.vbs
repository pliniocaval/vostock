'Insira os dados conforme necessidade.
dim StartIP, EndIP

StartIP = inputbox("DIGITE O INICIO DO RANGE")
EndIP = inputbox("DIGITE O FIM DO RANGE")
Data = Date()
DtAt = Split(data,"/",-1)
DataLog = DtAt(0)&DtAt(1)&DtAt(2)

TempFilename = "Log de instalação - " &StartIP& " - "&EndIP& " - " &DataLog& ".txt"

LF = chr(10)
const ForReading = 1, ForWriting = 2, ForAppending = 3

currentIP = StartIP

j = 2

'Inicia o LOG

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set Arquivo = objFSO.CreateTextFile(Tempfilename)
Set WSHShell = WScript.CreateObject("WScript.Shell")
arquivo.writeline "Iniciando processo"
Do

Strhost = currentIP
if Ping(strHost) = True Then
 On Error Resume Next
 Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & _
 strHost & "\root\cimv2")
 
  if Err.number <> 0 Then
  
    Erro = Err.Description
   Err.Clear
     arquivo.writeline "IP " & strHost & " - " & StrTerminal & " - Terminal encontrado, retornado Erro: "&Erro
   Else

    
   Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
   For Each objComputer in colSettings
   if objcomputer.name = " " Then
   StrTerminal = "Off Line"
   else
   StrTerminal = objcomputer.name
   end if
   StrUsuario = objcomputer.username
   
'progamas a serem instalados
EsetU = "c:\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\" & StrTerminal & " c:\suporte\inst\ESETUninstaller.exe /product=nodv34 /force /nosafemode /reboot"
EsetU2 = "c:\pstools\psexec.exe -d \\" & StrTerminal & " c:\suporte\inst\ESETUninstaller.exe /product=nodv34 /force /nosafemode"   
remove = "c:\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\" & StrTerminal & " c:\ti\suporte\reparar.bat"
mxm = "c:\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\" & StrTerminal & " \\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\vbs\mxm.vbs"
time = "c:\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\" & StrTerminal & " net time \\csrv01 /set /y"
GPUP = "c:\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\" & StrTerminal & " gpupdate"
robo = "c:\pstools\pskill.exe -u cemusa\informatica -p 654321 -d \\" & StrTerminal & " robocopy.exe "
psshut = "c:\PsTools\psshutdown.exe -f -r -u cemusa\informatica -p 654321 \\" & StrTerminal
   'Executa a instalação
   Set objFSO = CreateObject("Scripting.FileSystemObject")
    'objFSO.DeleteFile "\\"&StrTerminal&"\c$\logs\cemusa.log", True
    'WSHShell.Run EsetU, 0, false
	'WSHShell.Run EsetU, 0, false
	'WSHShell.Run remove, 0, False
	'WSHShell.Run mxm, 0, False
	'WSHShell.Run time, 0, False
	'WSHShell.Run GPUP, 0, False
	'WSHShell.Run robo, 0, False
	WSHShell.Run psshut, 0, False
   'Grava Log
     arquivo.writeline "IP " & strHost & " - "&StrTerminal&" - "&Strusuario&" - Processo iniciada - " & Time()
   
   

  Next
  
  End If


 Else
  
  arquivo.writeline "Terminal " & strHost & " não pode ser encontrado "
 
end if
 xx = currentIP
 currentIP = newIP(xx)
  j = j+1
Loop Until currentIP = EndIP
arquivo.writeline "Processo Finalizado"
Arquivo.close()

 '--------------------------------
'Função Ping via WMI
'--------------------------------
Function Ping(strHost)

  dim objPing, objRetStatus

  set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
   ("select * from Win32_PingStatus where address = '" & strHost & "'")

  for each objRetStatus in objPing
    if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
  Ping = False
      'WScript.Echo "Status code is " & objRetStatus.StatusCode
    else
      Ping = True
      'Wscript.Echo "Bytes = " & vbTab & objRetStatus.BufferSize
      'Wscript.Echo "Time (ms) = " & vbTab & objRetStatus.ResponseTime
      'Wscript.Echo "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
    end if
  next
 
End Function

'------------
'function for increasing the IP number
'------------
function newip(xx)
   dim n,n1,n2,n3,n4,v1,v2,v3,v4
   n = 1
   n0 = n

   while mid (xx,n,1) <> "."
     n = n+1
   wend

   n1 = n
   n = n+1

   while mid (xx,n,1) <> "."
     n = n+1
   wend

   n2 = n
   n = n+1

   while mid (xx,n,1) <> "."
     n = n+1
   wend

   n3 = n
   n4 = len(xx)
   v1 = mid (xx,n0,n1-1)
   v2 = mid (xx,n1+1,n2-n1-1)
   v3 = mid (xx,n2+1,n3-n2-1)
   v4 = mid (xx,n3+1,n4-n3)
   v4 = v4+1

   if v4 > 255 then
     v3 = v3+1
     v4 = 0
   end if

   if v3 > 255 then
     v2 = v2+1
     v3 = 0
     v4 = 0
   end if

   if v2 > 255 then
     v1 = v1+1
     v2 = 0
     v3 = 0
     v4 = 0
   end if

   return = (v1 & "." & v2 & "." & v3 & "." & v4)
   newIP = return
end function

'------------
'function for validating the IP address
'------------
function ValidIP(xx)

   dim n,n0,n1,n2,n3,n4,v1,v2,v3,v4,s,s1,s2,s3,s4
   n = 1
   n0 = n
   s = 1
   return = "valid"

   s1 = InStr(s, xx, ".", 1)
   s2 = InStr(s1+1, xx, ".", 1)
   s3 = InStr(s2+1, xx, ".", 1)
   s4 = len(xx)+1

   if s1-s < 1 then
     return = "invalid"
   elseif s1-s > 3 then
     return = "invalid"
   elseif s2-s1 < 1 then
     return = "invalid"
   elseif s2-s1 > 4 then
     return = "invalid"
   elseif s3-s2 < 1 then
     return = "invalid"
   elseif s3-s2 > 4 then
     return = "invalid"
   elseif s4-s3 < 1 then
     return = "invalid"
   elseif s4-s3 > 4 then
     return = "invalid"
   else
     while mid (xx,n,1) <> "."
        n = n+1
     wend

     n1 = n
     n = n+1

     while mid (xx,n,1) <> "."
        n = n+1
     wend

     n2 = n
     n = n+1

     while mid (xx,n,1) <> "."
        n = n+1
     wend

     n3 = n
     n4 = len(xx)
     v1 = mid (xx,n0,n1-1)
     v2 = mid (xx,n1+1,n2-n1-1)
     v3 = mid (xx,n2+1,n3-n2-1)
     v4 = mid (xx,n3+1,n4-n3)

     if v4 > 255 then
        return = "invalid"
     end if

     if v3 > 255 then
        return = "invalid"
     end if

     if v2 > 255 then
        return = "invalid"
     end if

     if v1 > 255 then
        return = "invalid"
     end if
   end if

   ValidIP = return

end function