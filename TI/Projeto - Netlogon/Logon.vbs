'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 21/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes
  
'msgbox "Remover drivers mapeados"
'Set colDrives = objNet.EnumNetworkDrives
'For i = 0 to colDrives.Count-1 Step 2
'    objNet.RemoveNetworkDrive colDrives.Item(i), true, true
'Next

'msgbox "Alterando Registro"
'msgbox "Alterando Registro - USB"
const HKEY_LOCAL_MACHINE = &H80000002
'objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start",4 ,"REG_DWORD"
'objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Modem\Start",4 ,"REG_DWORD"
objShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\googletalk"
'objShell.Run ("netsh firewall set opmode disable"),0 , False

'msgbox "Criando pastas"
If Not objFso.FolderExists(LOGUSER) Then objFso.CreateFolder(LOGUSER)
If Not objFso.FolderExists(outlook) Then objFso.CreateFolder(outlook)
If Not objFso.FolderExists(outlookuser) Then objFso.CreateFolder(outlookuser)
If Not objFso.FolderExists(outlookbkp) Then objFso.CreateFolder(outlookbkp)
If Not objFso.FolderExists(outlookbkpuser) Then objFso.CreateFolder(outlookbkpuser)
If Not objFso.FolderExists(ti) Then objFso.CreateFolder(ti)
If Not objFso.FolderExists(suploc) Then objFso.CreateFolder(suploc)
If Not objFso.FolderExists(htaloc) Then objFso.CreateFolder(htaloc)
If Not objFso.FolderExists(instloc) Then objFso.CreateFolder(instloc)
If Not objFso.FolderExists(LOGLOC) Then objFso.CreateFolder(LOGLOC)
If Not objFso.FolderExists(uninstloc) Then objFso.CreateFolder(uninstloc)
If Not objFso.FolderExists(locmxm) Then objFso.CreateFolder(locmxm)

'msgbox "Limpa Logs com mais de 5MB
Set objFolder = objFSO.GetFolder(LOGUSER)
Set colFiles = objFolder.Files
For Each objFile in colFiles
if objFile.Size >= 5242880 Then
objFSO.DeleteFile(LOGUSER & "\" & objFile.Name)
end if
Next

'msgbox "Tela de Logon"
If Not objFso.FileExists(htaloc&"\Logon.hta") Then objFSO.CopyFile scripts & "\hta\Logon.hta" , htaloc&"\Logon.hta", OverwriteExisting
objShell.Run (htaloc&"\logon.hta"),1 ,False

'MsgBox "Segurança"
'objShell.Run (scripts&"\vbs\sec.vbs"),0 , True

'msgbox "Copia de arquivos"
objShell.Run (scripts&"\vbs\copia.vbs"),0 , True

'msgbox "assinatura de email"
objShell.Run (scripts&"\vbs\ass.vbs"),0 , False

'msgbox "Critica de saida"
if left(ucase(computador),3)="IMA" then wscript.quit
if left(ucase(computador),4)="CSRV" then wscript.quit
if left(ucase(computador),6)="CBSB04" then wscript.quit
if left(ucase(computador),7)="SQLSCPI" then wscript.quit
if left(ucase(computador),7)="CEMUSA-" then wscript.quit
if left(ucase(computador),3)="MXM" then wscript.quit

Set atalhoLnk = objShell.CreateShortcut(desktop & "\Reparar Rede.lnk")
atalhoLnk.TargetPath = suploc&"\Reparo.bat"
atalhoLnk.Description = "Reparo da rede"
atalhoLnk.WorkingDirectory = lochta
atalhoLnk.WindowStyle = 1
atalhoLnk.IconLocation = htaloc &"\img\logo.ico"
atalhoLnk.Save

'msgbox "BGinfo"
objFSO.DeleteFile USERPROFILE & "\AppData\Local\Temp\bginfo.bmp"
objFSO.DeleteFile USERPROFILE & "\Configurações locais\Temp\bginfo.bmp"
Wscript.Sleep 2500
objShell.Run BgInfo,0 , False 

'msgbox "Recadastramento dos usuario "
objShell.Run (scripts&"\vbs\cad.vbs"),0 , True

'msgbox "Inventario da Estação"
objShell.Run (scripts&"\vbs\inventario.vbs"),0 , False

'msgbox "Impressoras"
objShell.Run (scripts&"\vbs\print.vbs"),0 , False

'msgbox "instalações"
'objShell.Run (scripts&"\vbs\uninstall.vbs"),0 , True

'msgbox "instalações"
'objShell.Run (scripts&"\vbs\install.vbs"),0 , True

'msgbox "instalações"
objShell.Run (scripts&"\vbs\audit.vbs"),0 , False

'msgbox "fim"
wscript.quit