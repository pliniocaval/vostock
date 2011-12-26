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

'variaveis
atalho = objShell.SpecialFolders("Desktop")

mxm = "\\csrv02\mx-manager\RDP\"
locmxm = "c:\mxm\"
logs = "c:\logs\"

'Apagar o log se for maior que 10MB
If objFSO.FileExists(logs&"mxm.log") Then
set file = objFSO.GetFile(logs&"mxm.log")
  if file.Size >= 10485760 Then
    objFSO.DeleteFile(logs&"mxm.log")
  End If
End If

'diretorios
objFSO.CreateFolder locmxm
objFSO.CreateFolder logs

'função de copya
robo = "c:\suporte\robocopy.exe "& mxm &" "& locmxm &" /MIR /TEE /LOG+:" & logs &"mxm.log"

computador = objNet.ComputerName
strTS = "CSRV04"
'WScript.Echo computador
if UCASE(computador) = strTs Then
atalho= objShell.SpecialFolders("Desktop")
objnet.RemoveNetworkDrive "T:", true, true
objnet.MapNetworkDrive "T:", "\\csrv02\MX-Manager"
objFSO.DeleteFile atalho&"\MXM.lnk"
objFSO.DeleteFile atalho&"\Microsoft Office O*.lnk"

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
set file = objFSO.GetFile(logs &"mxm.log")		
If DateDiff("d", file.DateLastModified, Now) > 4 Then 
'wscript.echo File
'wscript.echo File.DateLastModified
objShell.Run robo, 0, True
End If
objFSO.CopyFile locmxm&"mxm.lnk", atalho&"\MXM.lnk", True
wscript.quit
End if



