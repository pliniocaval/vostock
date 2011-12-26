'Script do logon
'autoria Leonardo Vivas
'Versão 0.2
'criação 03/06/2009
'modificação 03/06/2009
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Não parar em caso de erros
On Error Resume Next

'Proxy
configproxy = "10.10.1.9:3128"
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", configproxy, "REG_SZ"

	 
Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run"
strKeyPath1 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
strStringValueName = "Adobe Reader Speed Launcher"
strStringValueName1 = "Adobe ARM"
strStringValueName2 = "SunJavaUpdateSched"

'64bits
oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath,strStringValueName
oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath,strStringValueName1	 
oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath,strStringValueName2

'32bits
oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath1,strStringValueName
oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath1,strStringValueName1
oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath1,strStringValueName2	 
	 

'horario de verão

strTimeServer = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\DisableAutoDaylightTimeSet"
strTimeServer2 = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\DynamicDaylightTimeDisabled"
objShell.RegWrite strTimeServer,1,"REG_DWORD"
objShell.RegWrite strTimeServer2,1,"REG_DWORD"
'Wscript.Echo "horario ok"

'Blank
objShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\blank\"

'Proxy
configproxy = "10.10.1.9:3128"
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", configproxy, "REG_SZ"

