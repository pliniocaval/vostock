'Script do logon
'autoria Leonardo Vivas
'Versão 0.5
'criação 03/06/2009
'modificação 17/12/2010
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

' Não parar em caso de erros
'On Error Resume Next

'Variaveis
Dim objnet, objShell, objFSO, objSysInfo
Dim strComp, strDom, StrUser
Dim sDepartment, sLocation, sComputerName, sDomain, sGroups, objUser
Dim sStatus, intSeconds, sDesktop, sScriptDir, iTimerID

'Levantando informaçoes
strComp = UCase((objNet.ComputerName))
strDom = UCase((objNet.UserDomain))
strUser = objNet.UserName
sUserDN = objSysInfo.UserName
'separa a localização e o departamento atraves da OU. 
arrDept = split(sUserDN, ",")
sDepartment = mid(arrDept(1), 4) 
sLocation = mid(arrDept(2), 4)
									

' Remover drivers mapeados
Set colDrives = objNet.EnumNetworkDrives
For i = 0 to colDrives.Count-1 Step 2
    objNet.RemoveNetworkDrive colDrives.Item(i), true, true
Next

									
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser &  ",user")
For Each objGroup In objUser.Groups
'Mapeamento por Grupo / Locaion / Departamento baseado na Unidade Organizacional
If (ucase(objGroup.Name) = "USUÁRIOS DO DOMÍNIO") or (ucase(objGroup.Name) = "DOMAIN USERS") Then
objnet.MapNetworkDrive "G:" , "\\cemusadobrasil.com.br\Geral"
objnet.MapNetworkDrive "M:", "\\cemusadobrasil.com.br\departamentos"
objnet.MapNetworkDrive "P:", "\\cemusadobrasil.com.br\pdcontas"
If Not objFSO.FolderExists("\\cemusadobrasil.com.br\user$\" & sLocation & "\" & strUser) Then objFSO.CreateFolder("\\cemusadobrasil.com.br\user$\" & sLocation & "\" & strUser) 					
objnet.MapNetworkDrive "U:", "\\cemusadobrasil.com.br\user$\" & sLocation & "\" & strUser
objnet.MapNetworkDrive "H:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\" & sDepartment
end if

'Ações por Departamento baseado na Unidade Organizacional
'******COPACABANA******
'***INFORMATICA***
If (ucase(sDepartment) = "INFORMATICA") <> 0 Then
	If (ucase(objGroup.Name) = "SUPORTE") <> 0 Then
	objnet.MapNetworkDrive "X:", "\\csrv06\TI$"
	End If
End If
'***COMERCIAL***
If (ucase(sDepartment) = "COMERCIAL") <> 0 Then
	If (ucase(objGroup.Name) = "VENDAS") <> 0 Then
	objnet.MapNetworkDrive "V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas"
	End If
End If
'***COMERCIAL***
If (ucase(sDepartment) = "JURIDICO") <> 0 Then
	If (ucase(objGroup.Name) = "VENDAS") <> 0 Then
	objnet.MapNetworkDrive "V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas"
	End If
	If (ucase(objGroup.Name) = "COMERCIAL") <> 0 Then
	objnet.MapNetworkDrive "I:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial"
	End If
	If (ucase(objGroup.Name) = "SECRETARIAS") <> 0 Then
	objnet.MapNetworkDrive "J:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\secretarias"
	End If
	If (ucase(objGroup.Name) = "DIRETORIA") <> 0 Then
	objnet.MapNetworkDrive "K:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\diretoria"
	End If
End If
'***DIRETORIA***
If (ucase(sDepartment) = "DIRETORIA") <> 0 Then
	If (ucase(objGroup.Name) = "VENDAS") <> 0 Then
	objnet.MapNetworkDrive "V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas"
	End If
	If (ucase(objGroup.Name) = "COMERCIAL") <> 0 Then
	objnet.MapNetworkDrive "I:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial"
	End If
	If (ucase(objGroup.Name) = "CONTABILIDADE") <> 0 Then
	objnet.MapNetworkDrive "J:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\contabilidade"
	End If
End If
'******SÃO CRISTOVÃO******
'***RH***
If (ucase(sDepartment) = "RH") <> 0 Then
	If (ucase(objGroup.Name) = "RH") <> 0 Then
	objnet.MapNetworkDrive "X:", "\\sqlscpi\bomark"
	End If
End If
' Ações por Grupo
'BLOQUEIOS
If (ucase(objGroup.Name) = "GERENTES") Then
objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start",3 ,"REG_DWORD"
objShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Modem\Start",3 ,"REG_DWORD"
End If
'FOTOS
If (ucase(objGroup.Name) = "CIRCUITOS") Then
objnet.MapNetworkDrive "O:", "\\csrv06\Circuitos-Fotos"
End If
'FTP
If (ucase(objGroup.Name) = "FTP") Then
objnet.MapNetworkDrive "S:", "\\10.10.1.2\ftp"
End If
'Acesso MXM
If (ucase(objGroup.Name) = "MXM-REMOTO") Then
objShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\todos\mxm.vbs"), 0, True
End If
next								

computador = objNet.ComputerName
if left(ucase(computador),4)="CSRV" then wscript.quit
if left(ucase(computador),3)="IMA" then wscript.quit
if left(ucase(computador),4)="VIRU" then wscript.quit
if left(ucase(computador),2)="TS" then wscript.quit
'wscript.echo computador
objShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\todos\cad.vbs")
objShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\todos\reg.vbs")
objShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\todos\log.vbs"), 0, True

wscript.quit									
									
									