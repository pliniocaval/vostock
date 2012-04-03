'Script de Startup | Leonardo Vivas
' -----------------------------------------------------------------'

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'msgbox "Não parar em caso de erros"
On Error Resume Next

'MsgBox "Capturando Diretorio do Script"
DIRS = oFSO.GetParentFolderName(WScript.ScriptFullName)

'msgbox "Carregando Variaveis Remotas"
Loadfile(DIRS & "\SYS\DIRL.INI")

'msgbox "Carregando variaveis"
Loadfile(DIRS & "\SYS\VAR.INI")

'MsgBox "Limpa Versão anterior do Script"
'ApagaRaiz(TI)
oShell.Run ("cmd.exe /C rmdir /s /q" & " " & TI),0 , True

'msgbox "Criando pastas"
CriaPasta(TI)
CriaPasta(TIATU)
CriaPasta(HTA)
CriaPasta(IMG)
CriaPasta(PROGS)
CriaPasta(LOGS)
CriaPasta(SYS)
CriaPasta(PARA)
CriaPasta(SUPORTE)
CriaPasta(USERLOGS)

'MsgBox "sincroniza arquivos"
If oFso.FileExists("C:\Windows\SysWOW64\RoboCopy.exe") Then
Else
If oFso.FileExists("C:\Windows\System32\RoboCopy.exe") Then
Else
If oFso.FileExists(PROGS & "\RoboCopy.exe") Then
Else
CopiaArquivo DIRS & "\PROGS\RoboCopy.exe",PROGS & "\RoboCopy.exe"
Robo = PROGS & "\RoboCopy.exe"
End If
End If
End If

oShell.Run Robo & " " & DIRS & "\ " & TIATU & "\ " & RoboOPSYNC & LOGS & "\STARTUP.log", 0, True

Function CriaPasta(pasta)
 If Not oFso.FolderExists(pasta) Then oFso.CreateFolder(pasta)
 End Function
 
Function Loadfile(File)
  varfile = File
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE
End Function

Function ApagaRaiz(Pasta)
set folder = oFSO.getFolder (Pasta)
if folder.Subfolders.count > 0 then
for each SubFolder in folder.Subfolders
ApagaRaiz SubFolder
SubFolder.delete
next
end if
for each file in folder.files
set objFile = oFSO.GetFile(file)
objFile.attributes = 0
File.delete
next
if folder.Subfolders.count = 0 and folder.files.count=0 and Folder.Path=Pasta then
Folder.delete true
end if
End Function

Function CopiaContPasta(origem,destino)
Set objFolder = oFSO.GetFolder(origem)
Set colFiles = objFolder.Files
For Each objFile in colFiles
oFSO.CopyFile (origem & "\" & objFile.Name),  (destino & "\" & objFile.Name), OverwriteExisting
Next
End Function

Function CopiaArquivo(Origem,Destino)
Do Until Carquivo = True
	If not oFSO.FileExists(Destino) Then
	Carquivo = False
	oFSO.CopyFile Origem , Destino, OverwriteExisting
	wscript.sleep 2500
	Else
	set FOrigem = oFSO.getFile (Origem)
	set FDestino = oFSO.getFile (Destino)
	If FOrigem.dateLastModified > FDestino.dateLastModified Then
	oFSO.CopyFile Origem , Destino, OverwriteExisting
	End If
	Carquivo = True
	End If
 Loop
End Function


