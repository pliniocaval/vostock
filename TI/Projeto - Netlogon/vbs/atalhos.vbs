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

Set objUser = GetObject("LDAP://" & sUserDN)

'Informação da UO
arrDept = split(sUserDN, ",")
sLocation = mid(arrDept(2), 4)

'MsgBox "Atalho para Reparo de Rede"  
Set ReparoLnk = objShell.CreateShortcut(desktop & "\Reparar Rede.lnk")
ReparoLnk.TargetPath = suploc&"\Reparo.bat"
ReparoLnk.Description = "Reparo da rede"
ReparoLnk.WorkingDirectory = lochta
ReparoLnk.WindowStyle = 1
ReparoLnk.IconLocation = htaloc &"\img\logo.ico"
ReparoLnk.Save

'msgbox "Create shortcut to Desktop."
Set DepLnk = objShell.CreateShortcut(Desktop & "\Departamentos " & sLocation & ".lnk")
DepLnk.TargetPath = "M:\" & sLocation & "\"
DepLnk.Description = "Atalho para " & sLocation
DepLnk.WorkingDirectory = "M:\"
DepLnk.WindowStyle = 1
DepLnk.Save

'MsgBox "Atalho para Outlook"
'objFSO.DeleteFile Desktop&"\Microsoft Office O*.lnk"
'objFSO.DeleteFile USERPROFILE & "\Dados de aplicativos\Microsoft\Internet Explorer\Quick Launch\Microsoft Office O*.lnk"
'objFSO.DeleteFile vAPPDATA & "\Microsoft\Internet Explorer\Quick Launch\Microsoft Office O*.lnk"
'objFSO.CopyFile "\\csrv06\ti$\office\atalhos\Microsoft Office Outlook.lnk" ,Desktop & "\Microsoft Office Outlook.lnk"
'objFSO.CopyFile "\\csrv06\ti$\office\atalhos\Microsoft Office Outlook.lnk" ,vAPPDATA & "\Microsoft\Internet Explorer\Quick Launch\Microsoft Office Outlook.lnk"