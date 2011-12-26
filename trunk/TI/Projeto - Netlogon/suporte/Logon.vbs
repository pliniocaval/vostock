'Script do logon
'autoria Leonardo Vivas
'Versão 0.3
'criação 03/06/2009
'modificação 12/07/2010
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Não parar em caso de erros
On Error Resume Next

'variaveis
user = "\\cemusadobrasil.com.br\user$\"
scripts ="\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\"

computador = objNet.ComputerName
if left(ucase(computador),4)="CSRV" then
'wscript.echo computador
else
if left(ucase(computador),3)="IMA" then
'wscript.echo computador
else
if left(ucase(computador),4)="VIRU" then
'wscript.echo computador
else
'Ass Email
'objShell.Run (scripts&"todos\ass.vbs")
'objShell.Run (scripts&"todos\usblog.vbs")
end if
end if
end if

' Remover drivers mapeados
'Set colDrives = objNet.EnumNetworkDrives
'For i = 0 to colDrives.Count-1 Step 2
'    objNet.RemoveNetworkDrive colDrives.Item(i), true, true
'Next

''''Mapeamentos Basicos'''''
' Mapemento de Todos:

objnet.MapNetworkDrive "G:" , "\\cemusadobrasil.com.br\Geral"
objnet.MapNetworkDrive "M:", "\\cemusadobrasil.com.br\departamentos"

''''Fim dos Mapeamentos Basicos'''''

''''Mapeamentos por Grupos''''
strDom = objNet.UserDomain
strUser = objNet.UserName
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser &  ",user")

For Each objGroup In objUser.Groups



'Mapeamentos de Localidade 
If objGroup.Name = "Copacabana" Then
 If objFSO.FolderExists(user&"Copacabana\"& objnet.UserName) Then
 Set objFolder = objFSO.GetFolder(user&"Copacabana\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Copacabana\"& objnet.UserName
 'Wscript.Echo "Pasta Existe"
 Else
 'Wscript.Echo "Pasta não Existe"
 Set objFolder = objFSO.CreateFolder(user&"Copacabana\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Copacabana\"& objnet.UserName
  End If
  End If
  
If objGroup.Name = "SC" Then
 If objFSO.FolderExists(user&"São Cristovão\"& objnet.UserName) Then
 Set objFolder = objFSO.GetFolder(user&"São Cristovão\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"São Cristovão\"& objnet.UserName
  'Wscript.Echo "Pasta Existe"
 Else
 'Wscript.Echo "Pasta não Existe"
 Set objFolder = objFSO.CreateFolder(user&"São Cristovão\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"São Cristovão\"& objnet.UserName
  End If
 	End If
	
If objGroup.Name = "Brasilia" Then
 If objFSO.FolderExists(user&"Brasilia\"& objnet.UserName) Then
 Set objFolder = objFSO.GetFolder(user&"Brasilia\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Brasilia\"& objnet.UserName 
  'Wscript.Echo "Pasta Existe"
 Else
 'Wscript.Echo "Pasta não Existe"
 Set objFolder = objFSO.CreateFolder(user&"Brasilia\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Brasilia\"& objnet.UserName 
  End If
  End If

If objGroup.Name = "Salvador" Then
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 If objFSO.FolderExists(user&"Salvador\"& objnet.UserName) Then
 Set objFolder = objFSO.GetFolder(user&"Salvador\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Salvador\"& objnet.UserName 
 'Wscript.Echo "Pasta Existe"
 Else
 'Wscript.Echo "Pasta não Existe"
 Set objFolder = objFSO.CreateFolder(user&"Salvador\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Salvador\"& objnet.UserName 
 End If
 End If
 
If objGroup.Name = "Manaus" Then
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 If objFSO.FolderExists(user&"Manaus\"& objnet.UserName) Then
 Set objFolder = objFSO.GetFolder(user&"Manaus\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Manaus\"& objnet.UserName 
 'Wscript.Echo "Pasta Existe"
 Else
 'Wscript.Echo "Pasta não Existe"
 Set objFolder = objFSO.CreateFolder(user&"Manaus\"& objnet.UserName)
 objnet.MapNetworkDrive "U:", user&"Manaus\"& objnet.UserName 
 End If
 End If
 
If objGroup.Name = "SP" Then

  Set objFSO = CreateObject("Scripting.FileSystemObject")
 If objFSO.FolderExists(user&"São Paulo\"& objnet.UserName) Then
 Set objFolder = objFSO.GetFolder(user&"São Paulo\"& objnet.UserName)
	objnet.MapNetworkDrive "U:", user&"São Paulo\"& objnet.UserName 
  'Wscript.Echo "Pasta Existe"
 Else
 'Wscript.Echo "Pasta não Existe"
 Set objFolder = objFSO.CreateFolder(user&"São Paulo\"& objnet.UserName)
	objnet.MapNetworkDrive "U:", user&"São Paulo\"& objnet.UserName
  End If
  End If

'Fim dos Mapeamentos de Localidade
' Mapeamentos por grupo

If objGroup.Name = "FTP" Then
 objnet.MapNetworkDrive "S:", "\\10.10.1.2\ftp"
 End If 

If objGroup.Name = "Circuitos" Then
    objnet.MapNetworkDrive "O:", "\\csrv06\Circuitos-Fotos"
	End If

'Drivers Copacabana
If objGroup.Name = "Suporte" Then
   objnet.MapNetworkDrive "X:", "\\csrv06\TI$"
   End If
   
If objGroup.Name = "TI" Then
   objnet.MapNetworkDrive "H:", "\\csrv01\Copacabana$\Informatica"
   End If

If objGroup.Name = "Contabilidade" Then
 objnet.MapNetworkDrive "H:", "\\csrv01\Copacabana$\Contabilidade"
  	End If

If objGroup.Name = "Secretarias" Then
 objnet.MapNetworkDrive "T:", "\\cemusadobrasil.com.br\departamentos\Copacabana\Secretarias"
 objnet.MapNetworkDrive "I:", "\\csrv01\Diretoria"
	End If

If objGroup.Name = "Comercial" Then
 objnet.MapNetworkDrive "H:", "\\csrv01\Copacabana$\Comercial"
 objnet.MapNetworkDrive "V:", "\\csrv01\Copacabana$\Comercial\vendas"
  End If
 
'Drive SC
If objGroup.Name = "Producao" Then
 objnet.MapNetworkDrive "H:", "\\sqlscpi\departamentos\São Cristovão\Produção"
    End If

If objGroup.Name = "RH" Then
    objnet.MapNetworkDrive "H:", "\\sqlscpi\departamentos\São Cristovão\RH"
    objnet.MapNetworkDrive "X:", "\\sqlscpi\bomark"   
	End If

If objGroup.Name = "Compras" Then
    objnet.MapNetworkDrive "X:", "\\sqlscpi\Compras"
	End If

If objGroup.Name = "SCPI" Then
	objnet.MapNetworkDrive "H:", "\\sqlscpi\SCPI2"
	objnet.MapNetworkDrive "T:", "\\sqlscpi\SCPI"
	End If
'Divers SP
If objGroup.Name = "SP" Then
 objnet.MapNetworkDrive "H:", "\\cemusa-srv\cemusa"
 End If
 
'Drive BSB
If objGroup.Name = "ProducaoBSB" Then
   objNet.MapNetworkDrive "H:", "\\cbsb04\Publico"
   End If

'Driver SSA
If objGroup.Name = "ProducaoSSA" Then
    End If

'Drive MN
If objGroup.Name = "ProducaoMAN" Then
	End If

If objGroup.Name = "Diretoria" Then
 objnet.MapNetworkDrive "I:", "\\csrv01\Diretoria"
 objnet.MapNetworkDrive "H:", "\\csrv01\Copacabana$\Contabilidade"
 objnet.MapNetworkDrive "T:", "\\csrv01\Copacabana$\Comercial"
 objnet.MapNetworkDrive "V:", "\\csrv01\Copacabana$\Comercial\vendas"
 End If
 
 'Acesso MXM
If objGroup.Name = "mxm-remoto" Then
objShell.Run (scripts&"todos\mxm.vbs"), 0, True
	End If
Next
wscript.quit
	