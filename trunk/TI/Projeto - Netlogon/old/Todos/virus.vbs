'Script do logon
'autoria Leonardo Vivas
'Versão 0.2
'criação 03/06/2009
'modificação 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

suploc="C:\suporte\"
LOGUSER="c:\logs\"
uphlog="cemusa.log"

objFSO.DeleteFile "c:\logs\cemusa.log", True

If objFSO.FileExists(LOGUSER&uphlog) Then 
 Set objFolder = objFSO.GetFile(LOGUSER&uphlog)
 'WScript.Echo suploc&uphlog
 wscript.quit
 Else 

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
virus64 = "c:\suporte\inst\cemusa64.exe /qb! REBOOT=" & Chr(34) & "ReallySuppress" & Chr(34)
objshell.run virus64, 0, True
'objFSO.DeleteFile suploc&"cemusa32.exe"
objFSO.CopyFile suploc&"cemusa64.exe" , suploc&"inst\cemusa64.exe", True
objFSO.CopyFile suploc&"cemusa32.exe" , suploc&"inst\cemusa32.exe", True
objFSO.CopyFile suploc&"cemusa64.exe" , LOGUSER&uphlog, True
else
'msgbox "32b"
virus32 = "c:\suporte\inst\cemusa32.exe /qb! REBOOT=" & Chr(34) & "ReallySuppress" & Chr(34)
objshell.run  virus32, 0, True
'objFSO.DeleteFile suploc&"cemusa64.exe"
objFSO.CopyFile suploc&"cemusa64.exe" , suploc&"inst\cemusa64.exe", True
objFSO.CopyFile suploc&"cemusa32.exe" , suploc&"inst\cemusa32.exe", True
objFSO.CopyFile suploc&"cemusa32.exe" , LOGUSER&uphlog, True

end if
end if


