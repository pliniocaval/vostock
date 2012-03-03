Set objShell = CreateObject("WScript.Shell")
WScript.Echo objShell.CurrentDirectory
objShell.CurrentDirectory = "C:\Temp"
WScript.Echo objShell.CurrentDirectory