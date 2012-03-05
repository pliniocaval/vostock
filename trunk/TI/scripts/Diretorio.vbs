Set oShell = CreateObject("Wscript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")

For i = 0 to Ubound(arrPath) - 1
    DIR = DIR & arrPath(i) & "\"
Next 

oShell.CurrentDirectory = DIR

Wscript.Echo oFSO.GetParentFolderName(WScript.ScriptFullName) & VbCrLf & oShell.CurrentDirectory