'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 14/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Não parar em caso de erros
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

'Apagar impressoras inuteis
Set objWMIService = GetObject("winmgmts:\\" & computador & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where DeviceID = 'Microsoft XPS Document Writer'")

For Each objPrinter in colInstalledPrinters
    objPrinter.Delete_
Next

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where DeviceID = 'Microsoft Shared Fax Driver'")

For Each objPrinter in colInstalledPrinters
    objPrinter.Delete_
Next

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where DeviceID = 'Fax'")

For Each objPrinter in colInstalledPrinters
    objPrinter.Delete_
Next

wscript.quit