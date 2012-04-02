'Script Para reparo | Leonardo Vivas
' ----------------------------------------------

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Captura e volta 1 nivel do diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")
For i = 0 to Ubound(arrPath) - 1
    DIRS = DIRS & arrPath(i) & "\"
Next 
oShell.CurrentDirectory = DIRS

'msgbox "Não parar em caso de erros"
'On Error Resume Next

'msgbox "Carregando Variaveis Remotas"
DIRLfile = DIRS & "\SYS\DIRL.INI"
  Set DIRL = oFSO.OpenTextFile(DIRLfile)
  DIRLFILE =   DIRL.ReadAll
  DIRL.close
  execute DIRLFILE

'msgbox "Carregando Variaveis Locais"

varfile = SYS & "\VAR.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'msgbox "Carregando Arquivo de Funções"
varfile = SYS & "\FNC.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE

'msgbox "Carregando Arquivo de Parametrização"
varfile = SYS & "\PARA\EMP.INI"
  Set VAR = oFSO.OpenTextFile(varfile)
  VARFILE =   VAR.ReadAll
  VAR.close
  execute VARFILE
  
	Dim arrDept
	
	Set objSysInfo = CreateObject("ADSystemInfo")
	sDN = objSysInfo.DomainDNSName
	sUserDN = objSysInfo.UserName
	Set objUser = GetObject("LDAP://" & sDN & "/" & sUserDN)
	
	'Busca informação do usuario e do computador.
	sUserName = oNet.UserName
	sComputerName = UCase((oNet.ComputerName))
	sDomain = UCase((oNet.UserDomain))
	sDisplayName = trim(objUser.DisplayName)
	
	'busca grupos do usuario
	sGroups = GetGroups(sUserDN)
	
	'Captura a UO do usuario. (assumindo que os usuarios estão assim definidos: Dominio->Localidade->Departmento->Usuario/Grupos)
	arrDept = split(sUserDN, ",")
	sDepartment = mid(arrDept(1), 4) 'Definir a profundidade. 
									'EX: CN=UserName,OU=Users,OU=Departmento,OU=Localização,DC=seu,DC=dominio,DC=com,DC=br; arrDept(1) = OU=Departmento
	sLocation = mid(arrDept(2), 4) 'Set number in array where department OU name is found. 
									'EX: CN=UserName,OU=Users,OU=Departmento,OU=Localização,DC=seu,DC=dominio,DC=com,DC=br; arrDept(2) = OU=Localização								
	
	'Se não conseguir o nome completo use o nome de usuario.
	If sDisplayName = "" Then

		sDisplayName = sUserName
	End If
	
	Err.Clear
	
If InStr(ucase(sGroups),"USUÁRIOS DO DOMÍNIO") or InStr(ucase(sGroups),"DOMAIN USERS") <> 0 Then 

	If oFSO.DriveExists("U:") Then
			ShowStat("U: Já Existe")
		Else
			If Not oFSO.FolderExists("\\" & sDN & "\user$\" & sLocation & "\" & sUserName) Then oFSO.CreateFolder("\\" & sDN & "\user$\" & sLocation & "\" & sUserName)
			If Not MapDrive("U:", "\\" & sDN & "\user$\" & sLocation & "\" & sUserName) Then
				If Not MapDrive("U:", "\\10.10.1.4\user$\" & sLocation & "\" & sUserName) Then

		   		Else

		   		End If
		   	Else

		  	End If
		End If
		
	If oFSO.DriveExists("G:") Then
			ShowStat("G: Já Existe")
		Else
			If Not MapDrive("G:", "\\" & sDN & "\Geral") Then
				If Not MapDrive("G:", "\\10.10.1.1\Geral") Then

		   		Else

		   		End If
		   	Else

		  	End If
		End If
		
	If oFSO.DriveExists("M:") Then
			ShowStat("M: Já Existe")
		Else
			If Not MapDrive("M:", "\\" & sDN & "\Departamentos") Then
				If Not MapDrive("M:", "\\10.10.1.4\Departamentos") Then

		   		Else

		   		End If
		   	Else

		  	End If
	End If
	
End If

	If oFSO.DriveExists("H:") Then
			ShowStat("H: Já Existe")
		Else
			If Not MapDrive("H:", "\\" & sDN & "\Departamentos\" & sLocation & "\" & sDepartment) Then
				If Not MapDrive("H:", "\\10.10.1.4\Departamentos\" & sLocation & "\" & sDepartment) Then
		  			
		   		Else

		   		End If
		   	Else

		  	End If
	End If
	
If oFSO.DriveExists("P:") Then
			ShowStat("P: Já Existe")
			Else
			If Not MapDrive("P:", "\\csrv01\PDContas") Then 
		 		If Not MapDrive("P:", "\\10.10.1.1\PDContas") Then 

		    	Else

		   		End If
		   	Else

		  	End If
End If

	
'---------------------gRUPOS-------------------------

If InStr(ucase(sGroups),"CIRCUITOS") <> 0 Then 
	If oFSO.DriveExists("O:") Then
			ShowStat("O: Já Existe")
			Else
			If Not MapDrive("O:", "\\cemusadobrasil.com.br\departamentos\Circuitos-Fotos") Then 
		 		If Not MapDrive("O:", "\\10.10.1.8\Circuitos-Fotos") Then 

		    	Else

		   		End If
		   	Else

		  	End If
		End If
				
End IF

If InStr(ucase(sGroups),"FTP") <> 0 Then 
	If oFSO.DriveExists("S:") Then
			ShowStat("S: Já Existe")
			Else
			If Not MapDrive("S:", "\\csrv05\ftp") Then 
		 		If Not MapDrive("S:", "\\10.10.1.2\ftp") Then 

		    	Else

		   		End If
		   	Else

		  	End If
		End If
				
End IF

If InStr(ucase(sDepartment),"INFORMATICA") <> 0 Then 
		If InStr(ucase(sGroups),"SUPORTE") <> 0 Then
		If oFSO.DriveExists("X:") Then

			Else
			If Not MapDrive("X:", "\\csrv06\TI$") Then
		 		If Not MapDrive("X:", "\\10.10.1.8\TI$") Then 

		    	Else

		   		End If
		   	Else

		  	End If
		End If
		End If		
End IF

If InStr(ucase(sDepartment),"SCPI") <> 0 Then 
		If InStr(ucase(sGroups),"COMPRAS") <> 0 Then
		If oFSO.DriveExists("I:") Then

			Else
			If Not MapDrive("I", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\Compras") Then 
		 		If Not MapDrive("I:", "\\10.10.2.5\departamentos\" & sLocation & "\Compras") Then 

		    	Else

		   		End If
		   	Else

		  	End If
		End If
		End If		
End IF

If InStr(ucase(sDepartment),"RH") <> 0 Then 
		If InStr(ucase(sGroups),"RH") <> 0 Then
		If oFSO.DriveExists("X:") Then

			Else
			If Not MapDrive("X", "\\SQLSCPI\BOMARK") Then 
		 		If Not MapDrive("X:", "\\10.10.2.5\BOMARK") Then 

		    	Else

		   		End If
		   	Else

		  	End If
		End If
		End If		
End IF

If InStr(ucase(sLocation),"SÃO PAULO") <> 0 Then
		If InStr(ucase(sGroups),"VENDAS") <> 0 Then 
		If oFSO.DriveExists("V:") Then

			Else
			If Not MapDrive("V:", "\\cemusadobrasil.com.br\Departamentos\Copacabana\Vendas") Then 
		 		If Not MapDrive("V:", "\\10.10.1.4\Departamentos\Copacabana\Vendas") Then 

		    	Else

		   		End If
		   	Else

		  	End If
		End If
		End If		
End IF

'---------------------------OUTROS------------------------------

'If left(ucase(sComputerName),4)="MXM-" then
'	If InStr(ucase(sGroups),"DIRETORIA") <> 0 Then
'	oShell.Run ("\\10.10.1.2\logon\VBS\03-AssEmail.vbs"),0 , False
	'oShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\vbs\outlook.vbs"), 0, True
	'oShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\vbs\outlookbkp.vbs"), 0, False
'	End IF
'End IF

If InStr(ucase(sGroups),"MXM-REMOTO") <> 0 Then
	oShell.Run ("\\10.10.1.2\logon\VBS\04-MXM.vbs"), 0, True
End IF

'If InStr(ucase(sGroups),"GERENTES") <> 0 Then
'	const HKEY_LOCAL_MACHINE = &H80000002
'	oShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start",3 ,"REG_DWORD"
'	oShell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Modem\Start",3 ,"REG_DWORD"

'End If

wscript.quit

Function GetGroups(sUDN)
	On Error Resume Next
	
	'Function to return user's Group Memberships
	Set objUser2 = GetObject("LDAP://" & sDN & "/" & sUDN)
	
	If objUser2.primaryGroupID = 513 Then
		sList = sList & "Domain Users" & VbCrLf
	Else 
		If objUser2.primaryGroupID = 512 Then
			sList = sList & "Domain Admins" & VbCrLf
		End If
	End If

	oMemberOf = objUser2.GetEx("memberOf")

	For Each oGroup In oMemberOf
		oGroup = Mid(oGroup, 4, 330)
		arrGroup = Split(oGroup, ",")
		sList = sList & arrGroup(0) & VbCrLf
	Next 
	
	Set objUser2 = Nothing
	
	GetGroups = sList
End Function

Function MapDrive(strDrive, strShare)
  On Error Resume Next
  Err.Clear
  If oFso.DriveExists(strDrive) Then
    Set objDrive = oFso.GetDrive(strDrive)
    If Err.Number <> 0 Then
      Err.Clear
      MapDrive = False
      Exit Function
    End If
    If CBool(objDrive.DriveType = 3) Then
      oNet.RemoveNetworkDrive strDrive, True, True
    Else
      MapDrive = False
      Exit Function
    End If
    Set objDrive = Nothing
  End If
  oNet.MapNetworkDrive strDrive, strShare
  If Err.Number = 0 Then
    MapDrive = True
  Else
    Err.Clear
    MapDrive = False
  End If
  On Error GoTo 0
End Function