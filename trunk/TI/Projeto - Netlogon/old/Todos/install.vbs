'Script do logon
'autoria Leonardo Vivas
'Versão 0.2
'criação 03/06/2009
'modificação 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

suploc="C:\suporte\inst\"
LOGUSER="c:\logs\"
esetlog = "cemusa.log"

'objFSO.DeleteFile "c:\logs\cemusa.log", True

If objFSO.FileExists(LOGUSER&esetlog) Then 
 Set objFolder = objFSO.GetFile(LOGUSER&esetlog)
 'WScript.Echo suploc&uphlog
 wscript.quit
 Else 

 lRet = 2
Do While lRet = 2
   Msg = VbCrLf
   Msg = Msg & "Seu AntiVirus precisa ser atualizado." & chr(10) & VbCrLf
   Msg = Msg & "Favor pressionar pressionar OK para iniciar a instalação." & chr(10)& VbCrLf
   Msg = Msg & "Favor Não executar nenhum programa ate a instalação ser concluida." & chr(10)& VbCrLf
   Msg = Msg & "Ao final desta voce sera informado." & chr(10) 
   
lRet  =   MsgBox(msg,0,"Cemusa Informa")
Loop
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
strProperties = "TotalPhysicalMemory, UserName, SystemType, Description, DaylightInEffect"
objClass = "Win32_ComputerSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colSys = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colSys
PC_Type = objItem.SystemType
next

if left(ucase(PC_Type),3)="X64" then 
virus64 = "c:\suporte\inst\cemusa64.exe /qn REBOOT=" & Chr(34) & "ReallySuppress" & Chr(34)
'objshell.run virus64, 0, True
eset64
Wscript.Sleep 60000 
objFSO.CopyFile suploc&"cemusa64.exe" , LOGUSER&esetlog, True
else
'msgbox "32b"
virus32 = "c:\suporte\inst\cemusa32.exe /qn REBOOT=" & Chr(34) & "ReallySuppress" & Chr(34)
eset32
Wscript.Sleep 60000 
'objshell.run  virus32, 0, True
objFSO.CopyFile suploc&"cemusa32.exe" , LOGUSER&esetlog, True

end if
end if

 lRet = 2
Do While lRet = 2
   Msg = VbCrLf
   Msg = Msg & "Seu AntiVirus foi atualizado." & chr(10) & VbCrLf
   Msg = Msg & "Obrigado por sua Atenção." & chr(10) 
   
lRet  =   MsgBox(msg,0,"Cemusa Informa")
Loop


Function eset32 
Set WshShell = CreateObject("Wscript.Shell") 
Set WshEnv = WshShell.Environment("PRocess") 
WshShell.Run "runas.exe /user:" & "cemusa\informatica" & " " & Chr(34) & "c:\suporte\inst\cemusa32.exe /qn REBOOT=" & Chr(34) & "ReallySuppress" & Chr(34) & Chr(34)
Wscript.Sleep 800 
WshShell.AppActivate WshEnv("SystemRoot") & "\system32\runas.exe" 
Wscript.Sleep 200 
WshShell.SendKeys "654321" & "~" 
Wscript.Sleep 5000 
Set WshShell = Nothing 
Set WshEn = Nothing 
End Function 

Function eset64
Set WshShell = CreateObject("Wscript.Shell") 
Set WshEnv = WshShell.Environment("PRocess") 
WshShell.Run "runas.exe /user:" & "cemusa\informatica" & " " & Chr(34) & "c:\suporte\inst\cemusa64.exe /qn REBOOT=" & Chr(34) & "ReallySuppress" & Chr(34) & Chr(34)
Wscript.Sleep 800 
WshShell.AppActivate WshEnv("SystemRoot") & "\system32\runas.exe" 
Wscript.Sleep 200 
WshShell.SendKeys "654321" & "~" 
Wscript.Sleep 5000 
Set WshShell = Nothing 
Set WshEn = Nothing 
End Function 