<!--
  Logon Script - logon.hta 
  Written by Geary Epperson.
  
-->

<head>
<title></title>

<script language="VBScript">
	'Prevent Window flickering on load.
	Me.ResizeTo 500,300
	'Move Window off screen.
    Me.MoveTo ((Screen.Width)),((Screen.Height))
</script>

<HTA:APPLICATION
     APPLICATIONNAME="LogonScript"
     BORDER="thin"
     BorderStyle="complex"
     SCROLL="no"
     maximizebutton="no"
  	 minimizebutton="no"  	 
  	 SINGLEINSTANCE="yes"
     WINDOWSTATE="normal"
     SysMenu="no"
     ContextMenu="no"
     Icon='logon.ico'
>
</head>

<script language="VBScript">

Dim FSO, oShell, oNetwork, objSysInfo, sUserDN, objUser
Dim sDepartment, sLocation, sUserName, sComputerName, sDomain, sDisplayName, sGroups, sDN
Dim sStatus, intSeconds, sDesktop, sScriptDir, iTimerID

'Configure for your domain. See "MainScript" Sub for drive mappings.
sDN = "cemusadobrasil.com.br"

Sub Window_onLoad
	On Error Resume Next
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set oShell = CreateObject("WScript.Shell")
	Set oNetwork = CreateObject("WScript.Network")

    'Get User's information.
    UserInfo
	
    'User's Desktop for deploying shortcuts.
    sDesktop = oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\"
    
    'Populate Window with user info.
	document.title = sDomain & " Logon Script - " & sDepartment 'Changes title bar to reflect domain and department the current user is logging onto.
	DisplayName.InnerHTML = sDisplayName
	UserName.InnerHTML = sUserName
	ComputerName.InnerHTML = sComputerName
	
	'Replace logo with Dept logo if found.
	HTA = location.pathname
	HTA_Path = Left(HTA,InStrRev(HTA,"\"))
	
	'Replace logo with Dept logo if found. Place department logo in same dir as logon.hta. Name should be logo-department.jpg, where department is the OU name.
	If FSO.FileExists(HTA_Path & "IMG\logo-" & sDepartment & ".jpg") Then
		Logo.src = "IMG\logo-" & sDepartment & ".jpg"
	Else
		Logo.src = "IMG\logo-default.jpg"
	End If
	
	'Move to top left of screen.
	Me.MoveTo 10,10
	
	'Run Main Logon Script
	MainScript
	
	'Countdown timer before closing. Set time in seconds.
	intSeconds = 10
	iTimerID = window.setInterval("Count", 1000)
End Sub

Sub Default_Buttons
    If Window.Event.KeyCode = 13 Then
    End If
End Sub

Sub UserInfo
	On Error Resume Next
	
	Dim arrDept
	
	Set objSysInfo = CreateObject("ADSystemInfo")
	
	sUserDN = objSysInfo.UserName
	Set objUser = GetObject("LDAP://" & sDN & "/" & sUserDN)
	
	'Find User and Computer info.
	sUserName = oNetwork.UserName
	sComputerName = UCase((oNetwork.ComputerName))
	sDomain = UCase((oNetwork.UserDomain))
	sDisplayName = trim(objUser.DisplayName)
	
	'Find Group Memberships
	sGroups = GetGroups(sUserDN)
	
	'Get department name from DN. (Assuming users OU in AD is setup as Domain->Department->Users->UserObject)
	arrDept = split(sUserDN, ",")
	sDepartment = mid(arrDept(1), 4) 'Set number in array where department OU name is found. 
									'ie: CN=UserName,OU=Users,OU=Department,DC=your,DC=domain,DC=com; arrDept(2) = OU=Department
	sLocation = mid(arrDept(2), 4) 'Set number in array where department OU name is found. 
									'ie: CN=UserName,OU=Users,OU=Department,DC=your,DC=domain,DC=com; arrDept(2) = OU=Department								
	
	'If Full Name isn't found, set as username.
	If sDisplayName = "" Then

		sDisplayName = sUserName
	End If
	
	Err.Clear
	
	Set objSysInfo = Nothing
	Set objUser = Nothing
End Sub

Sub MainScript

' Remover drivers mapeados
Set colDrives = oNetwork.EnumNetworkDrives
For i = 0 to colDrives.Count-1 Step 2
    oNetwork.RemoveNetworkDrive colDrives.Item(i), true, true
Next

	'On Error Resume Next
	
	' *********************************
	' ***	 Mapeamentos Comuns  	***
	' *********************************
	If InStr(ucase(sGroups),"USUÁRIOS DO DOMÍNIO") or InStr(ucase(sGroups),"DOMAIN USERS") <> 0 Then 'Domain Users get H and P mappings.
		If FSO.DriveExists("P:") Then
			ShowStat("P: Já Existe")
			Else
			If Not MapDrive("P:", "\\cemusadobrasil.com.br\pdcontas") Then 
		 		If Not MapDrive("P:", "\\10.10.1.1\pdcontas") Then 
		    		ShowStat("P: - Falha no Mapeamento")
		    	Else
		    		ShowStat("P: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("P: - Mapeado com Sucesso")
		  	End If
		End If
		
		If FSO.DriveExists("G:") Then
			ShowStat("G: Já Existe")
		Else
			If Not MapDrive("G:", "\\cemusadobrasil.com.br\Geral") Then
				If Not MapDrive("G:", "\\10.10.1.1\Geral") Then
		   			ShowStat("G: - Falha no Mapeamento")
		   		Else
		   			ShowStat("G: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("G: - Mapeado com Sucesso")
		  	End If
		End If
		
		If FSO.DriveExists("M:") Then
			ShowStat("M: Já Existe")
		Else
			If Not MapDrive("M:", "\\cemusadobrasil.com.br\departamentos") Then
				If Not MapDrive("M:", "\\10.10.1.1\departamentos") Then
		   			ShowStat("M: - Falha no Mapeamento")
		   		Else
		   			ShowStat("M: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("M: - Mapeado com Sucesso")
		  	End If
		End If
		
	If FSO.DriveExists("U:") Then
			ShowStat("U: Já Existe")
		Else
			If Not FSO.FolderExists("\\cemusadobrasil.com.br\user$\" & sLocation & "\" & sUserName) Then FSO.CreateFolder("\\cemusadobrasil.com.br\user$\" & sLocation & "\" & sUserName) 
			If Not MapDrive("U:", "\\cemusadobrasil.com.br\user$\" & sLocation & "\" & sUserName) Then
				If Not MapDrive("U:", "\\10.10.1.1\Geral") Then
		   			ShowStat("U: - Falha no Mapeamento")
		   		Else
		   			ShowStat("U: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("U: - Mapeado com Sucesso")
		  	End If
		End If
		
		If FSO.DriveExists("H:") Then
			ShowStat("H: Já Existe")
			Else
			If Not MapDrive("H:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\" & sDepartment) Then 
		 		If Not MapDrive("H:", "\\10.10.1.1\departamentos\" & sLocation & "\" & sDepartment) Then 
		    		ShowStat("H: - Falha no Mapeamento")
		    	Else
		    		ShowStat("H: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("H: - Mapeado com Sucesso")
		  	End If
		End If
		
	End If

	' *********************************
	' *** Fim dos Mapeamentos Comuns***
	' *********************************
	
    ' *********************************
	' ******** RECURSOS GERAIS ********
	' *********************************
	If InStr(ucase(sGroups),"CIRCUITOS") <> 0 Then 
	If FSO.DriveExists("O:") Then
			ShowStat("O: Já Existe")
			Else
			If Not MapDrive("O:", "\\csrv06\Circuitos-Fotos") Then 
		 		If Not MapDrive("O:", "\\10.10.1.8\Circuitos-Fotos") Then 
		    		ShowStat("O: - Falha no Mapeamento")
		    	Else
		    		ShowStat("O: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("O: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	If InStr(ucase(sGroups),"FTP") <> 0 Then 
	If FSO.DriveExists("S:") Then
			ShowStat("S: Já Existe")
			Else
			If Not MapDrive("S:", "\\csrv06\Circuitos-Fotos") Then 
		 		If Not MapDrive("S:", "\\10.10.1.8\Circuitos-Fotos") Then 
		    		ShowStat("S: - Falha no Mapeamento")
		    	Else
		    		ShowStat("S: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("S: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	If InStr(ucase(sGroups),"MXM-REMOTO") <> 0 Then
	oShell.Run ("\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\todos\mxm.vbs"), 0, True
	ShowStat("MXM - Disponibilizado com Sucesso")
	End IF
    ' *********************************
	' ***** FIM RECURSOS GERAIS *******
	' *********************************
	
	' *********************************
	' ***  MAPEAMENTOS POR GRUPOS	***
	' *********************************
	' ************* COPACABANA *************
	' ************* INFORMATICA *************
	If InStr(ucase(sDepartment),"INFORMATICA") <> InStr(ucase(sGroups),"SUPORTE") Then 
		If FSO.DriveExists("X:") Then
			ShowStat("X: Já Existe")
			Else
			If Not MapDrive("X:", "\\csrv06\TI$") Then
		 		If Not MapDrive("X:", "\\10.10.1.8\TI$") Then 
		    		ShowStat("X: - Falha no Mapeamento")
		    	Else
		    		ShowStat("X: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("X: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	' ************* CONTABILIDADE *************
	
	' ************* COMERCIAL *************
	If InStr(ucase(sDepartment),"COMERCIAL") <> InStr(ucase(sGroups),"VENDAS") Then 
		If FSO.DriveExists("V:") Then
			ShowStat("V: Já Existe")
			Else
			If Not MapDrive("V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas") Then 
		 		If Not MapDrive("V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas") Then 
		    		ShowStat("V: - Falha no Mapeamento")
		    	Else
		    		ShowStat("V: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("V: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	' ************* SECRTARIAS *************
	
	' ************* JURIDICO *************
	If InStr(ucase(sDepartment),"JURIDICO") <> InStr(ucase(sGroups),"VENDAS") Then 
		If FSO.DriveExists("V:") Then
			ShowStat("V: Já Existe")
			Else
			If Not MapDrive("V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas") Then 
		 		If Not MapDrive("V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas") Then 
		    		ShowStat("V: - Falha no Mapeamento")
		    	Else
		    		ShowStat("V: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("V: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	If InStr(ucase(sDepartment),"JURIDICO") <> InStr(ucase(sGroups),"DIRETORIA") Then 
		If FSO.DriveExists("I:") Then
			ShowStat("I: Já Existe")
			Else
			If Not MapDrive("I:", "\\cemusadobrasil.com.br\departamentos\diretoria") Then 
		 		If Not MapDrive("I:", "\\cemusadobrasil.com.br\departamentos\diretoria") Then 
		    		ShowStat("I: - Falha no Mapeamento")
		    	Else
		    		ShowStat("I: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("I: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	If InStr(ucase(sDepartment),"JURIDICO") and InStr(ucase(sGroups),"COMERCIAL") Then 
		If FSO.DriveExists("J:") Then
			ShowStat("J: Já Existe")
			Else
			If Not MapDrive("J:", "\\cemusadobrasil.com.br\departamentos\diretoria") Then 
		 		If Not MapDrive("J:", "\\cemusadobrasil.com.br\departamentos\diretoria") Then 
		    		ShowStat("J: - Falha no Mapeamento")
		    	Else
		    		ShowStat("J: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("J: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	' ************* DIRETORIA *************
	If InStr(ucase(sDepartment),"DIRETORIA") and InStr(ucase(sGroups),"VENDAS") Then 
		If FSO.DriveExists("V:") Then
			ShowStat("V: Já Existe")
			Else
			If Not MapDrive("V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas") Then 
		 		If Not MapDrive("V:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial\vendas") Then 
		    		ShowStat("V: - Falha no Mapeamento")
		    	Else
		    		ShowStat("V: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("V: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	If InStr(ucase(sDepartment),"DIRETORIA") and InStr(ucase(sGroups),"COMERCIAL") Then 
		If FSO.DriveExists("I:") Then
			ShowStat("I: Já Existe")
			Else
			If Not MapDrive("I:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial") Then 
		 		If Not MapDrive("I:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\comercial") Then 
		    		ShowStat("I: - Falha no Mapeamento")
		    	Else
		    		ShowStat("I: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("I: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	If InStr(ucase(sDepartment),"DIRETORIA") and InStr(ucase(sGroups),"CONTABILIDADE") Then 
		If FSO.DriveExists("J:") Then
			ShowStat("J: Já Existe")
			Else
			If Not MapDrive("J", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\contabilidade") Then 
		 		If Not MapDrive("J:", "\\cemusadobrasil.com.br\departamentos\" & sLocation & "\contabilidade") Then 
		    		ShowStat("J: - Falha no Mapeamento")
		    	Else
		    		ShowStat("J: - Mapeado com Sucesso")
		   		End If
		   	Else
		   		ShowStat("J: - Mapeado com Sucesso")
		  	End If
		End If
				
	End IF
	' ************* SÃO CRISTOVÃO *************
	' ************* SCPI *************
	' ************* COMPRAS *************
	' ************* RH *************
	
	' ************* BRASILIA *************
	
	' ************* SALVADOR *************
	
	' ************* MANAUS *************
	

	' *********************************
	' ***FIM DOS MEPAMENTOS POR GRUPO***
	' *********************************
	
	'Copy shortcut to Desktop.
	'FSO.CopyFile "\\server\share\Shortcut.lnk", sDesktop, False
End Sub

'-------------------- Functions --------------------------

Sub CloseSelf
	window.close
End Sub

Sub Hold
	document.all.lock.checked = True
	window.clearInterval(iTimerID)
	countdown.Style.Display = "none"
	btn_close.Style.Display = "inline"
End Sub

Sub Count
	'Bring script to front.
	window.focus()
	
	If intSeconds <> 0 Then
		countdown.InnerHTML = intSeconds
		intSeconds = intSeconds - 1
	Else
		If Not document.all.lock.checked Then
			CloseSelf
		End If
	End If
End Sub

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
  If FSO.DriveExists(strDrive) Then
    Set objDrive = FSO.GetDrive(strDrive)
    If Err.Number <> 0 Then
      Err.Clear
      MapDrive = False
      Exit Function
    End If
    If CBool(objDrive.DriveType = 3) Then
      oNetwork.RemoveNetworkDrive strDrive, True, True
    Else
      MapDrive = False
      Exit Function
    End If
    Set objDrive = Nothing
  End If
  oNetwork.MapNetworkDrive strDrive, strShare
  If Err.Number = 0 Then
    MapDrive = True
  Else
    Err.Clear
    MapDrive = False
  End If
  On Error GoTo 0
End Function

Function ShowStat(sMessage)
	sStatus = sMessage & VbCrLf & sStatus
	document.all.status.InnerText = sStatus
End Function

</script>

<body id="mainbody" bgcolor="white" style="font:Verdana; color:black" onclick="hold">
<style type="text/css">
.estilotextarea {background-color: transparent;border: 1px solid #000000;}


</style>
	<table width="100%" border="0" cellpadding="0">
		<tr valign="center">
			<td align="center" width="30%">
				<img name="Logo">					
			</td>			
			<td align="center" width="70%">
				<font size="3">Bem Vindo&nbsp;<strong><span style="color:blue" id="DisplayName"></span></strong>&nbsp;</font><br><br>
				<strong><font face="bold" size="3">Usuário:&nbsp;<span style="color:blue" id="UserName"></span>&nbsp;&nbsp;Computador:&nbsp;<span style="color:blue" id="ComputerName"></span></strong>
			</td>
		</tr>		
		<tr>
			<td>
			</td>		
		</tr>
	</table>
	<table width="100%" border="0" cellpadding="0">
		<tr align="left">
			<td>
				<textarea class=estilotextarea rows="9" name="status" cols="73" style="font-family: Verdana; font-weight:bold; font-size: 8pt"></textarea>
			</td>
		</tr>
		<hr color="red">	
	</table>	
	<table width="100%" border="0" cellpadding="0">
		<tr valign="top">
			<td align="left" width="50%">
				<font size="2.25">Manter esta janela aberta&nbsp;</font><input type="checkbox" name="lock">
			</td>
			<td align="right" width="50%">
				<span id="countdown"></span><input type="button" name="btn_close" style="display:none" value="Close" onclick="CloseSelf">
			</td>
		</tr>	
	</table>
</body>