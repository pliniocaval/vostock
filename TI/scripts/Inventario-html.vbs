':::::::::::::::::::::::::::::::::::::
':: ::
':: Configuração de Variáveis ::
':: ::
':::::::::::::::::::::::::::::::::::::
'Option Explicit
Dim strPC, strOpc, strHTML, objIE, Wrt_HTML, Var_Header_Title, Var_Header_Table_1
Dim Var_Digit_Box, MountHtml, strKey, strSubKey, objReg, arrSubKeys, strDisplayName, strDisplayVersion, strInstallLocation
Dim sComputer, oArgs, sOpt, sVisible, isValidParameters, Flag, strProperties, strDx, Var_Soft_Name, Var_Soft_Version
Dim Var_Soft_Build, Var_Msg_1, Var_Msg_2, VarMsg_Err1, VarMsg_Err2, VarMsg_Err3, LcId

Var_Soft_Name = ("Sistema de inventario da Cemusa do Brasil")
Var_Soft_Version = (" v.1.0.05")
Var_Soft_Build = (" r.2405")
Var_Msg_1 = ("Digite o NOME ou IP do Computador desejado:" & vbCrLf & "Para sair, digite 0 e confirme.")
Var_Msg_2 = ("Digite uma das opções abaixo:" & vbCrLf & vbCrLf & " 0 - Para sair." & vbCrLf & " 1 - Para somente visualizar." & vbCrLf & " 2 - Para visualizar e gerar um arquivo HTML." & vbCrLf & " 3 - Para somente gerar um arquivo em HTML." & vbCrLf)
VarMsg_Err1 = ("Você não digitou as informações solicitadas, a aplicação será finalizada.")
VarMsg_Err2 = ("Computador não existe. Favor digite corretamente as informações!")
VarMsg_Err3 = ("Não há ítens para serem exibidos.")
VarMsg_About = ("\n\nDesenvolvido por Leonardo Vivas")

':::::::::::::::::::::::::::::::::::::
':: ::
':: Configurações para o script ::
':: ::
':::::::::::::::::::::::::::::::::::::
LcId = 1046
Const wbemFlagReturnImmediately = 16
Const wbemFlagForwardOnly = 32
Const wbemCimtypeUint32 = 19
Const wbemCimtypeSint64 = 20
Const wbemCimtypeUint64 = 21
Const HKEY_LOCAL_MACHINE = &H80000002

':::::::::::::::::::::::::::::::::::::
':: ::
':: Configuração de filtro Dx ::
':: ::
':::::::::::::::::::::::::::::::::::::
strDx = True

':::::::::::::::::::::::::::::::::::::
':: ::
':: Sub Start IE e Monta HTML ::
':: ::
':::::::::::::::::::::::::::::::::::::
Sub StartIE(byval strHTML)
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE")
objIE.Navigate ("about:blank")
objIE.ToolBar = False
objIE.MenuBar = False
objIE.AddressBar = False
objIE.StatusBar = True
objIE.Resizable = False
objIE.Width = 700 
objIE.Height = 800 
objIE.Left = 0
objIE.Top = 0
objIE.Visible = strHTML
Do While (objIE.Busy) 
WScript.Sleep 100 
Loop

Set Wrt_HTML = objIE.Document 
Wrt_HTML.Open 
Wrt_HTML.Write HeaderHtml() & vbCrLf
MountHtml = MountHtml & "<div id='t_head'>" & Var_Soft_Name & Var_Soft_Version & Var_Soft_Build & "</div>" & vbCrLf
MountHtml = MountHtml & "<div id='l_head'>Início do inventário: " & AudData() & "</div>" & vbCrLf
MountHtml = MountHtml & "<div id='t_body'>" & vbCrLf

End Sub

':::::::::::::::::::::::::::::::::::::
':: ::
':: Sumário ::
':: ::
':::::::::::::::::::::::::::::::::::::
Sub DG_PCInfo(byval strPc)
On Error Resume Next

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strPC & "\root\cimv2")
Set objReg = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strPc & "\root\default:StdRegProv")

Trat_Err()
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Sumário do Sistema</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf

strProperties = "*"'"CSName, Caption, OSType, Version, OSProductSuite, BuildNumber, ProductType, OSLanguage, CSDVersion, InstallDate, RegisteredUser, Organization, SerialNumber, WindowsDirectory, SystemDirectory"
objClass = "Win32_OperatingSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colOS = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colOS
Host_Name = objItem.CSName
SO_Name = objItem.Caption
SO_Type = objItem.OSType
SO_Version = objItem.Version
SO_Suite = objItem.OSProductSuite
SO_Build = objItem.BuildNumber
SO_ProdType = objItem.ProductType
SO_Language = objItem.OSLanguage
SP_Version = objItem.CSDVersion
SO_InstDate = FormatDataTime(objItem.InstallDate)
SO_RegUser = objItem.RegisteredUser
SO_RegOrg = objItem.Organization
SO_Serial = objItem.SerialNumber
SO_WinDir = objItem.WindowsDirectory
SO_SysDir = objItem.SystemDirectory
If SO_Type = 16 Then
SO_Name = "Microsoft Windows 95"
ElseIf SO_Type = 17 Then
SO_Name = "Microsoft Windows 98"
End If
If SO_ProdType = 1 Then
SO_ProdType = "Estação de Trabalho"
ElseIf SO_ProdType = 2 Then
SO_ProdType = "Controlador de Domínio"
ElseIf SO_ProdType = 3 Then
SO_ProdType = "Servidor"
End If
If SO_Language = 1033 Then
SO_Language = "Inglês - Estados Unidos"
ElseIf SO_Language = 1046 Then
SO_Language = "Português - Brasil"
Else
SO_Language = "Outro idioma"
End If
If SO_Suite = 1 Then
SO_Suite = "Small Business"
ElseIf SO_Suite = 2 Then
SO_Suite = "Enterprise"
ElseIf SO_Suite = 4 Then
SO_Suite = "Backoffice"
ElseIf SO_Suite = 8 Then
SO_Suite = "Communication Server"
ElseIf SO_Suite = 16 Then
SO_Suite = "Terminal Server"
ElseIf SO_Suite = 18 Then
SO_Suite = "Enterprise e Terminal Server"
ElseIf SO_Suite = 32 Then
SO_Suite = "Small Business (Restrito)"
ElseIf SO_Suite = 64 Then
SO_Suite = "Embedded NT"
ElseIf SO_Suite = 128 Then
SO_Suite = "Data Center"
ElseIf SO_Suite = 256 Then
SO_Suite = "Single User"
ElseIf SO_Suite = 512 Then
SO_Suite = "Personal"
ElseIf SO_Suite = 1024 Then
SO_Suite = "Blade"
End If
MountHtml = MountHtml & "<li><span class='s_line' style='width: 195px;'>COMPUTADOR: </span><span class='s_line' style='width: 349px;'>" & CheckNull(UCase(Host_Name)) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Sistema Operacional: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão (Release): </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_Version) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Build: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_Build) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Função do Computador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_ProdType) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Idioma do Sistema: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_Language) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Service Pack: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SP_Version) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Suíte: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_Suite) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Data da instalação: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_InstDate) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Usuário registrado: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_RegUser) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Organização registrada: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_RegOrg) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Número serial: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_Serial) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Diretório do Windows: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_WinDir) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Diretório do sistema: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_SysDir) & "</span></li>" & vbCrLf
Next

strProperties = "TotalPhysicalMemory, UserName, SystemType, Description, DaylightInEffect"
objClass = "Win32_ComputerSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colSys = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colSys
PC_Mem = FormatValue(objItem.TotalPhysicalMemory)
PC_Logon = objItem.UserName
PC_Type = objItem.SystemType
PC_Info = objItem.Description
PC_HorVerao = objItem.DaylightInEffect
If IsNull(PC_Logon) Then
PC_Logon = "<span class='red'>Não há usuário logado neste sistema.</span>"
End If
If PC_HorVerao = True Then
PC_HorVerao = "<span class='red'>Desabilitado</span>"
Else
PC_HorVerao = "<span class='green'>Habilitado</span>"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Memória do sistema: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(PC_Mem) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Usuário logado: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(PC_Logon) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tipo do sistema: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(PC_Type) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Descrição: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(PC_Info) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Horário de Verão: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(PC_HorVerao) & "</span></li>" & vbCrLf
Next

Set objWMIServiceIE = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strPC & "\root\cimv2\Applications\MicrosoftIE")
strProperties = "Version"
objClass = "MicrosoftIE_FileVersion"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE file = 'iexplore.exe'"
Set colIESettings = objWMIServiceIE.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colIESettings
IE_Version = objItem.Version
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Internet Exploter: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(IE_Version) & "</span></li>" & vbCrLf
Next

strKey = "SOFTWARE\Microsoft\DirectX"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKey, arrSubKeys
For Each strSubKey In arrSubKeys
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKey & "\" & strSubKey, "Version", strDisplayDxVersion
DX_Version = strDisplayDxVersion
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>DirectX: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(DX_Version) & "</span></li>" & vbCrLf
Next

strKey = "SOFTWARE\Microsoft\MediaPlayer\PlayerUpgrade"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKey, arrSubKeys
For Each strSubKey In arrSubKeys
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKey & "\" & strSubKey, "PlayerVersion", strDisplayMPVersion
MP_Version = Replace(strDisplayMPVersion, ",", ".")
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Media Player: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MP_Version) & "</span></li>" & vbCrLf
Next

strProperties = "Description, MACAddress, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, DNSDomain, DNSDomainSuffixSearchOrder, DHCPEnabled, DHCPServer, WINSPrimaryServer, WINSSecondaryServer, ServiceName"
objClass = "Win32_NetworkAdapterConfiguration"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE IPEnabled = True AND ServiceName <> 'AsyncMac' AND ServiceName <> 'VMnetx' AND ServiceName <> 'VMnetadapter' AND ServiceName <> 'Rasl2tp' AND ServiceName <> 'PptpMiniport' AND ServiceName <> 'Raspti' AND ServiceName <> 'NDISWan' AND ServiceName <> 'RasPppoe' AND ServiceName <> 'NdisIP' AND ServiceName <> ''"
Set colAdapters = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colAdapters
Lan_Name = objItem.Description
Lan_Mac_Address = objItem.MACAddress
IP_Address = objItem.IPAddress
SubNet_Masc = objItem.IPSubnet
IP_Gateway = objItem.DefaultIPGateway
DNS_Server = objItem.DNSServerSearchOrder
DNS_Domain = objItem.DNSDomain
DNS_Domain_Sufix = objItem.DNSDomainSuffixSearchOrder
DHCP_Status = objItem.DHCPEnabled
DHCP_Server = objItem.DHCPServer
WINS_Server_1 = objItem.WINSPrimaryServer
WINS_Server_2 = objItem.WINSSecondaryServer
If DHCP_Status = True Then
DHCP_Status = "<span class='green'>Habilitado</span>"
Else
DHCP_Status = "<span class='red'>Desabilitado</span>"
End If
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Dispositivo de LAN: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Lan_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Endereço Mac</span>: <span class='li_itens' style='width: 349px;'>" & CheckNull(Lan_Mac_Address) & "</span></li>" & vbCrLf
If Not IsNull(IP_Address) Then
For i = 0 To UBound(IP_Address)
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Endereço IP: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(IP_Address(i)) & "</span></li>" & vbCrLf
Next 
End If
If Not IsNull(SubNet_Masc) Then
For i = 0 To UBound(SubNet_Masc)
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Máscara da Subnet: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SubNet_Masc(i)) & "</span></li>" & vbCrLf
Next 
End If
If Not IsNull(IP_Gateway) Then
For i = 0 To UBound(IP_Gateway)
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Servidor Gateway: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(IP_Gateway(i)) & "</span></li>" & vbCrLf
Next
End If
If Not IsNull(DNS_Server) Then
For i = 0 To UBound(DNS_Server)
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Servidor DNS: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(DNS_Server(i)) & "</span></li>" & vbCrLf
Next
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nome do Domínio: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(DNS_Domain) & "</span></li>" & vbCrLf

If Not IsNull(DNS_Domain_Sufix) Then
For i = 0 To UBound(DNS_Domain_Sufix)
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Sufixo DNS: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(DNS_Domain_Sufix(i)) & "</span></li>" & vbCrLf
Next 
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Status do DHCP: </span><span class='li_itens' style='width: 349px;'>" & DHCP_Status & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Servidor DHCP: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(DHCP_Server) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Servidor WINS Primário: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(WINS_Server_1) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Servidor WINS Secundário: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(WINS_Server_2) & "</span></li>" & vbCrLf
Next

MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Informações de Segurança ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Centro de Segurança</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>PROTEÇÃO ANTIVÍRUS: </span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

Set objWMIServiceAV = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strPC & "\root\SecurityCenter")
strProperties = "DisplayName, VersionNumber, CompanyName, OnAccessScanningEnabled, ProductUptoDate"
objClass = "AntiVirusProduct"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colSecurity = objWMIServiceAV.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colSecurity
AV_Name = objItem.DisplayName
AV_ProdVersion = objItem.VersionNumber
AV_Fab = objItem.CompanyName
AV_ScanStatus = objItem.OnAccessScanningEnabled
AV_ProdUpdate = objItem.ProductUptoDate
Next
If AV_ScanStatus = True Then
AV_ScanStatus = "<span class='green'>Ativado</span>"
Else
AV_ScanStatus = "<span class='red'>Desativado</span>"
End If
If AV_ProdUpdate = True Then
AV_ProdUpdate = "<span class='green'>OK</span>"
Else
AV_ProdUpdate = "<span class='red'>Produto não atualizado</span>.<br/>Atualize o antivírus o quanto antes."
End If
If AV_Name <> "" AND AV_ProdVersion <> "" AND AV_Fab <> "" AND AV_ScanStatus <> "" AND AV_ProdUpdate <> "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Antivírus: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(AV_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão do Antivírus: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(AV_ProdVersion) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(AV_Fab) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Escaneamento em tempo real: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(AV_ScanStatus) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Atualização das definições: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(AV_ProdUpdate) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Else
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'><span class='red'>Não é possível conectar ao Centro de Segurança</span>.<br/>Recurso apenas disponível no Windows XP SP2, Windows Server 2003 SP1 e Windows Media Center 2005.<br/>O Centro de Segurança do Windows não detecta todos os antivírus do mercado.</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If

strProperties = "DataExecutionPrevention_Available, DataExecutionPrevention_32BitApplications, DataExecutionPrevention_Drivers"
objClass = "Win32_OperatingSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colSecurity = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colSecurity
SO_DEPHard = objItem.DataExecutionPrevention_Available
SO_DEP32bitApp = objItem.DataExecutionPrevention_32BitApplications
SO_DEPDrivers = objItem.DataExecutionPrevention_Drivers
If SO_DEPHard = True Then
SO_DEPHard = "<span class='green'>Ativado</span>"
ElseIf SO_DEPHard = False Then
SO_DEPHard = "<span class='red'>Desativado</span><br/>Apenas processadores AMD Athlon 64, AMD Opteron e Intel Pentium 4 64-Bits em conjunto com o Windows XP SP2, Windows Server 2003 SP1 ou Windows Media Center 2005 possuem suporte ao recurso."
End If
If SO_DEP32bitApp = True Then
SO_DEP32bitApp = "<span class='green'>Ativado</span>"
Else
SO_DEP32bitApp = "<span class='red'>Desativado</span><br/>Apenas processadores AMD Athlon 64, AMD Opteron e Intel Pentium 4 64-Bits em conjunto com o Windows XP SP2, Windows Server 2003 SP1 ou Windows Media Center 2005 possuem suporte ao recurso."
End If
If SO_DEPDrivers = True Then
SO_DEPDrivers = "<span class='green'>Ativado</span>"
Else
SO_DEPDrivers = "<span class='red'>Desativado</span><br/>Apenas processadores AMD Athlon 64, AMD Opteron e Intel Pentium 4 64-Bits em conjunto com o Windows XP SP2, Windows Server 2003 SP1 ou Windows Media Center 2005 possuem suporte ao recurso."
End If
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>PROTEÇÃO SOFTWARE/HARDWARE: </span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Data Execution Prevention (DEP): </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_DEPHard) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>DEP em aplicativos de 32bits: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_DEP32bitApp) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>DEP em drivers: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(SO_DEPDrivers) & "</span></li>" & vbCrLf
Next

MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Informações da CPU ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Processador</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf

strProperties = "NumberOfProcessors"
objClass = "Win32_ComputerSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colCPU = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in ColCPU
CPU_Quant = objItem.NumberOfProcessors
MountHtml = MountHtml & "<li><span class='s_line' style='width: 195px;'>ÍTEM </span><span class='s_line' style='width: 349px;'>DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nº de Processadores: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Quant) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next

strProperties = "DeviceID, Name, Family, CurrentClockSpeed, MaxClockSpeed, LoadPercentage, ProcessorId, Availability, AddressWidth, Version, Revision, Stepping, PowerManagementSupported, CurrentVoltage, SocketDesignation, ExtClock, L2CacheSize, L2CacheSpeed, Manufacturer, Description"
objClass = "Win32_Processor"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colCPU = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colCPU
CPU_ID = objItem.DeviceID
CPU_Name = objItem.Name
CPU_Family = objItem.Family
CPU_Clock = FormatClock(objItem.CurrentClockSpeed)
CPU_Clock_Max = FormatClock(objItem.MaxClockSpeed)
CPU_Usage = FormatPerc(objItem.LoadPercentage)
CPU_CPUID = objItem.ProcessorId
CPU_Available = objItem.Availability
CPU_Address = FormatBit(objItem.AddressWidth)
CPU_Version = objItem.Version
CPU_Revision = objItem.Revision
CPU_Stepping = objItem.Stepping
CPU_PowManSup = objItem.PowerManagementSupported
CPU_CurrentVolt = FormatVolt(objItem.CurrentVoltage)
CPU_Socket = objItem.SocketDesignation
CPU_BUS = FormatClock(objItem.ExtClock)
CPU_CL2 = MemValue(objItem.L2CacheSize)
CPU_CL2Speed = FormatClock(objItem.L2CacheSpeed)
CPU_Manufacturer = objItem.Manufacturer
CPU_Info = objItem.Description
If CPU_ID = "CPU0" Then
CPU_ID = 1
ElseIf CPU_ID = "CPU1" Then
CPU_ID = 2
ElseIf CPU_ID = "CPU2" Then
CPU_ID = 3
ElseIf CPU_ID = "CPU3" Then
CPU_ID = 4
End If
If CPU_Family = 1 Then
CPU_Family = "Outra"
ElseIf CPU_Family = 2 Then
CPU_Family = "Não identificável"
ElseIf CPU_Family = 11 Then
CPU_Family = "Pentium® brand"
ElseIf CPU_Family = 12 Then
CPU_Family = "Pentium® Pro"
ElseIf CPU_Family = 13 Then
CPU_Family = "Pentium® II"
ElseIf CPU_Family = 14 Then
CPU_Family = "Pentium® processor with MMX technology"
ElseIf CPU_Family = 15 Then
CPU_Family = "Celeron?"
ElseIf CPU_Family = 16 Then
CPU_Family = "Pentium® II Xeon"
ElseIf CPU_Family = 17 Then
CPU_Family = "Pentium® III"
ElseIf CPU_Family = 18 Then
CPU_Family = "M1 Family"
ElseIf CPU_Family = 19 Then
CPU_Family = "M2 Family"
ElseIf CPU_Family = 24 Then
CPU_Family = "K5 Family"
ElseIf CPU_Family = 25 Then
CPU_Family = "K6 Family"
ElseIf CPU_Family = 26 Then
CPU_Family = "K6-2 Family"
ElseIf CPU_Family = 27 Then
CPU_Family = "K6-3 Family"
ElseIf CPU_Family = 28 Then
CPU_Family = "AMD Athlon? Processor Family"
'ElseIf CPU_Family = 29 Then
'CPU_Family = "AMD® Duron? Processor"
ElseIf CPU_Family = 31 Then
CPU_Family = "K6-2+ Family"
ElseIf CPU_Family = 120 Then
CPU_Family = "Crusoe? TM5000 Family"
ElseIf CPU_Family = 121 Then
CPU_Family = "Crusoe? TM3000 Family"
ElseIf CPU_Family = 130 Then
CPU_Family = "Itanium? Processor"
ElseIf CPU_Family = 176 Then
CPU_Family = "Pentium® III Xeon?"
ElseIf CPU_Family = 177 Then
CPU_Family = "Pentium® III Processor with Intel® SpeedStep? Technology"
ElseIf CPU_Family = 178 Then
CPU_Family = "Pentium® 4"
ElseIf CPU_Family = 179 Then
CPU_Family = "Intel® Xeon?"
ElseIf CPU_Family = 181 Then
CPU_Family = "Intel® Xeon? processor MP"
ElseIf CPU_Family = 182 Then
CPU_Family = "AMD AthlonXP? Family"
ElseIf CPU_Family = 183 Then
CPU_Family = "AMD AthlonMP? Family"
ElseIf CPU_Family = 184 Then
CPU_Family = "Intel® Itanium® 2"
ElseIf CPU_Family = 185 Then
CPU_Family = "AMD Opteron? Family"
ElseIf CPU_Family = 190 Then
CPU_Family = "K7 Family"
ElseIf CPU_Family = 300 Then
CPU_Family = "6x86 Family"
ElseIf CPU_Family = 301 Then
CPU_Family = "MediaGX Family"
ElseIf CPU_Family = 302 Then
CPU_Family = "MII Family"
ElseIf CPU_Family = 320 Then
CPU_Family = "WinChip Family"
End If
If CPU_Available = 1 Then
CPU_Available = "Outro"
ElseIf CPU_Available = 2 Then
CPU_Available = "Não Avaliável"
ElseIf CPU_Available = 3 Then
CPU_Available = "Executando/Full Power"
ElseIf CPU_Available = 4 Then
CPU_Available = "Perigo"
ElseIf CPU_Available = 5 Then
CPU_Available = "Em teste"
ElseIf CPU_Available = 6 Then
CPU_Available = "Não aplicável"
ElseIf CPU_Available = 7 Then
CPU_Available = "Desligado"
ElseIf CPU_Available = 8 Then
CPU_Available = "Off Line"
ElseIf CPU_Available = 9 Then
CPU_Available = "Off Duty"
ElseIf CPU_Available = 10 Then
CPU_Available = "Degradado"
ElseIf CPU_Available = 11 Then
CPU_Available = "Não instalado"
ElseIf CPU_Available = 12 Then
CPU_Available = "Erro de Instalação"
ElseIf CPU_Available = 13 Then
CPU_Available = "Power Save - Não Avaliável"
ElseIf CPU_Available = 14 Then
CPU_Available = "Power Save - Low Power Mode"
ElseIf CPU_Available = 15 Then
CPU_Available = "Power Save - Standby"
ElseIf CPU_Available = 16 Then
CPU_Available = "Power Cycle"
ElseIf CPU_Available = 17 Then
CPU_Available = "Power Save - Warning"
End If
If CPU_PowManSup = True Then
CPU_PowManSup = "Suportado"
Else
CPU_PowManSup = "Não suportado"
End If
If CPU_Manufacturer = "GenuineIntel" Then
CPU_Manufacturer = "Intel Corporation"
ElseIf CPU_Manufacturer = "AuthenticAMD" Then
CPU_Manufacturer = "AMD - Advanced Micro Devices, Inc."
End If
MountHtml = MountHtml & "<li><span class='s_line' style='width: 195px;'>PROCESSADOR Nº </span><span class='s_line' style='width: 349px;'>" & CheckNull(CPU_ID) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Processador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Família do Processador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Family) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Clock do Processador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Clock) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Clock máximo do Processador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Clock_Max) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Uso do Processador: </span><span class='li_itens' style='width: 349px;'>" & CPU_Usage & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Identificação da CPU: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_CPUID) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Status do Processador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Available) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Endereçamento da CPU: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Address) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Version) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Revisão: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Revision) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Stepping: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Stepping) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Gereciamento de energia: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_PowManSup) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Voltagem do processador: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_CurrentVolt) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Socket: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Socket) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Front Side Bus: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_BUS) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Cache L2: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_CL2) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Cache L2 clock: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_CL2Speed) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Informações sobre a CPU: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CPU_Info) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next

MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Informações de Memória ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Memória do sistema</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>INFORMAÇÕES DA MEMÓRIA DO SISTEMA</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "TotalPhysicalMemory, TotalPageFileSpace, TotalVirtualMemory, AvailableVirtualMemory"
objClass = "Win32_LogicalMemoryConfiguration"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colMemConfig = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in ColMemConfig
Mem_Physical = MemValue(objItem.TotalPhysicalMemory)
Mem_PageFile = MemValue(objItem.TotalPageFileSpace)
Mem_VM = MemValue(objItem.TotalVirtualMemory)
Mem_AvailableVM = MemValue(objItem.AvailableVirtualMemory)
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Memória física: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_Physical) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Arquivo de troca (swap): </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_PageFile) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Memória Virtual: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_VM) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Memória Virtual disponível: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_AvailableVM) & "</span></li>" & vbCrLf
Next
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Informações da Mainboard ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Placa-Mãe</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>INFORMAÇÕES SOBRE A PLACA-MÃE</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Product, Manufacturer, Model, OtherIdentifyingInfo, SerialNumber, PartNumber, Version"
objClass = "Win32_BaseBoard"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colMBoard = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colMBoard
MB_Product = objItem.Product
MB_Manufacturer = objItem.Manufacturer
MB_Model = objItem.Model
MB_NS = objItem.SerialNumber
MB_PartNumber = objItem.PartNumber
MB_Version = objItem.Version
MB_OtherInfo = objItem.OtherIdentifyingInfo
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Placa-Mãe: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_Product) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Modelo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_Model) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Número serial: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_NS) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Part Number: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_PartNumber) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_Version) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Outras informações: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(MB_OtherInfo) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If MB_Product = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If

strProperties = "Name, Manufacturer, BuildNumber, CurrentLanguage, ReleaseDate, SerialNumber, SMBIOSBIOSVersion, Version"
objClass = "Win32_BIOS"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colBios = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colBios
Bios_Name = objItem.Name
Bios_Manufacturer = objItem.Manufacturer
Bios_Build = objItem.BuildNumber
Bios_Lang = objItem.CurrentLanguage
Bios_ReleaseDate = FormatDataTime(objItem.ReleaseDate)
Bios_SN = objItem.SerialNumber
Bios_SMBiosVersion = objItem.SMBIOSBIOSVersion
Bios_Version = objItem.Version
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>BIOS</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Bios: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Bios Build: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_Build) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Idioma do Bios: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_Lang) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Data do Bios: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_ReleaseDate) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Número Serial: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_SN) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão SMBIOS: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_SMBiosVersion) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão do bios: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Bios_Version) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>SLOTS DE MEMÓRIA RAM</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "DeviceLocator, Capacity, DataWidth, FormFactor, MemoryType, HotSwappable, Manufacturer"
objClass = "Win32_PhysicalMemory"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE FormFactor = 6 OR FormFactor = 7 OR FormFactor = 8 OR FormFactor = 11 OR FormFactor = 12 OR FormFactor = 13"
Set colMEM = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in ColMEM
Mem_Bank = objItem.DeviceLocator
Mem_Size = FormatValue(objItem.Capacity)
Mem_Bits = FormatBit(objItem.DataWidth)
Mem_HotSwap = objItem.HotSwappable
Mem_FFactor = objItem.FormFactor
Mem_Fab = objItem.Manufacturer
Mem_Type = objItem.MemoryType
If Mem_FFactor = 0 Then
Mem_FFactor = "Não identificável"
ElseIf Mem_FFactor = 1 Then
Mem_FFactor = "Outro tipo"
ElseIf Mem_FFactor = 6 Then
Mem_FFactor = "Formato Proprietário"
ElseIf Mem_FFactor = 7 Then
Mem_FFactor = "SIMM"
ElseIf Mem_FFactor = 8 Then
Mem_FFactor = "DIMM"
ElseIf Mem_FFactor = 11 Then
Mem_FFactor = "RIMM"
ElseIf Mem_FFactor = 12 Then
Mem_FFactor = "SODIMM"
ElseIf Mem_FFactor = 13 Then
Mem_FFactor = "SRIMM"
End If
If Mem_Bank <> "" Then
Mem_Bank = "Slot " & Mem_Bank
End If
If Mem_Type = 0 Then
Mem_Type = "Não identificável"
ElseIf Mem_Type = 1 Then
Mem_Type = "Outro tipo"
ElseIf Mem_Type = 2 Then
Mem_Type = "DRAM"
ElseIf Mem_Type = 2 Then
Mem_Type = "DRAM Síncrona"
ElseIf Mem_Type = 4 Then
Mem_Type = "Cache DRAM"
ElseIf Mem_Type = 5 Then
Mem_Type = "EDO"
ElseIf Mem_Type = 6 Then
Mem_Type = "EDRAM"
ElseIf Mem_Type = 9 Then
Mem_Type = "RAM"
ElseIf Mem_Type = 11 Then
Mem_Type = "Flash"
ElseIf Mem_Type = 17 Then
Mem_Type = "SDRAM"
ElseIf Mem_Type = 19 Then
Mem_Type = "RDRAM"
ElseIf Mem_Type = 20 Then
Mem_Type = "DDR"
End If
If Mem_HotSwap = True Then
Mem_HotSwap = "Sim"
Else
Mem_HotSwap = "Não"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>" & CheckNull(Mem_Bank) & "</span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_Size) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Largura do endereçamento: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_Bits) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tipo do módulo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_FFactor) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tipo de memória: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_Type) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Hot-Swappable: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_HotSwap) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Mem_Fab) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Mem_Bank = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>CONTROLADORAS</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, Manufacturer"
objClass = "Win32_IDEController"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colCTRLIDE = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colCTRLIDE
CTRLIDE_Name = objItem.Name
CTRLIDE_Manufacturer = objItem.Manufacturer
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Controladora: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CTRLIDE_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CTRLIDE_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next

strProperties = "Name, Manufacturer"
objClass = "Win32_SCSIController"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colCTRLSCSI = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colCTRLSCSI
CTRLSCSI_Name = objItem.Name
CTRLSCSI_Manufacturer = objItem.Manufacturer
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Controladora: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CTRLSCSI_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CTRLSCSI_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Video ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Vídeo</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>PLACA DE VÍDEO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "DeviceID, Name, VideoProcessor, ProtocolSupported, VideoArchitecture, AdapterRAM, VideoMemoryType, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentBitsPerPixel, MinRefreshRate, MaxRefreshRate, DriverVersion, DriverDate"
objClass = "Win32_VideoController"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colVideo = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colVideo
VGA_DevID = objItem.DeviceID
VGA_Name = objItem.Name
VGA_GPU = objItem.VideoProcessor
VGA_Interface = objItem.ProtocolSupported
VGA_Arch = objItem.VideoArchitecture 
VGA_Ram = FormatValue(objItem.AdapterRAM)
VGA_Ram_Type = objItem.VideoMemoryType
VGA_BitPix = FormatBit(objItem.CurrentBitsPerPixel)
VGA_Resolution = objItem.CurrentHorizontalResolution & " x " & objItem.CurrentVerticalResolution & " Pixels"
VGA_RefrRate = FormatHz(objItem.MinRefreshRate & "~" & objItem.MaxRefreshRate)
VGA_Driver = objItem.DriverVersion
VGA_Driver_Date = FormatDataTime(objItem.DriverDate)
If VGA_DevID = "VideoController1" Then
VGA_DevID = "1"
ElseIf VGA_DevID = "VideoController2" Then
VGA_DevID = "2"
End If
If VGA_Interface = 1 Then
VGA_Interface = "Outro tipo"
ElseIf VGA_Interface = 2 Then
VGA_Interface = "Não identificável"
ElseIf VGA_Interface = 3 Then
VGA_Interface = "EISA"
ElseIf VGA_Interface = 4 Then
VGA_Interface = "ISA"
ElseIf VGA_Interface = 5 Then
VGA_Interface = "PCI"
ElseIf VGA_Interface = 14 Then
VGA_Interface = "VESA"
ElseIf VGA_Interface = 15 Then
VGA_Interface = "PCMCIA"
ElseIf VGA_Interface = 16 Then
VGA_Interface = "USB"
ElseIf VGA_Interface = 43 Then
VGA_Interface = "AGP"
Else
VGA_Interface = "-"
End If
If VGA_Arch = 1 Then
VGA_Arch = "Outro tipo"
ElseIf VGA_Arch = 2 Then
VGA_Arch = "Não identificável"
ElseIf VGA_Arch = 3 Then
VGA_Arch = "CGA"
ElseIf VGA_Arch = 4 Then
VGA_Arch = "EGA"
ElseIf VGA_Arch = 5 Then
VGA_Arch = "VGA"
ElseIf VGA_Arch = 6 Then
VGA_Arch = "SVGA"
ElseIf VGA_Arch = 7 Then
VGA_Arch = "MDA"
ElseIf VGA_Arch = 8 Then
VGA_Arch = "HGC"
ElseIf VGA_Arch = 9 Then
VGA_Arch = "MCGA"
ElseIf VGA_Arch = 10 Then
VGA_Arch = "8514A"
ElseIf VGA_Arch = 11 Then
VGA_Arch = "XGA"
ElseIf VGA_Arch = 12 Then
VGA_Arch = "Linear Frame Buffer"
ElseIf VGA_Arch = 160 Then
VGA_Arch = "PC-98"
Else
VGA_Arch = "-"
End If
If VGA_Ram_Type = 1 Then
VGA_Ram_Type = "Outro tipo"
ElseIf VGA_Ram_Type = 2 Then
VGA_Ram_Type = "Não identificável"
ElseIf VGA_Ram_Type = 3 Then
VGA_Ram_Type = "VRAM"
ElseIf VGA_Ram_Type = 4 Then
VGA_Ram_Type = "DRAM"
ElseIf VGA_Ram_Type = 5 Then
VGA_Ram_Type = "SRAM"
ElseIf VGA_Ram_Type = 6 Then
VGA_Ram_Type = "WRAM"
ElseIf VGA_Ram_Type = 7 Then
VGA_Ram_Type = "EDO RAM"
ElseIf VGA_Ram_Type = 8 Then
VGA_Ram_Type = "Burst Synchronous DRAM"
ElseIf VGA_Ram_Type = 9 Then
VGA_Ram_Type = "Pipelined Burst SRAM"
ElseIf VGA_Ram_Type = 10 Then
VGA_Ram_Type = "CDRAM"
ElseIf VGA_Ram_Type = 11 Then
VGA_Ram_Type = "3DRAM"
ElseIf VGA_Ram_Type = 12 Then
VGA_Ram_Type = "SDRAM"
ElseIf VGA_Ram_Type = 13 Then
VGA_Ram_Type = "SGRAM"
Else
VGA_Ram_Type = "-"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Controladora de vídeo nº </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_DevID) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Placa de vídeo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Processador gráfico: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_GPU) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Interface de vídeo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Interface) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tipo de vídeo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Arch) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Memória Ram: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Ram) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tipo de Memória Ram: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Ram_Type) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Resolução: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Resolution) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Pixels: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_BitPix) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Taxa de atualização: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_RefrRate) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Versão do driver: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Driver) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Data do Driver: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(VGA_Driver_date) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next

MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>MONITOR</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, MonitorManufacturer, MonitorType"
objClass = "Win32_DesktopMonitor"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colMonitor = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colMonitor
Monitor_Info = objItem.Name
Monitor_Fab = objItem.MonitorManufacturer
Monitor_Type = objItem.MonitorType
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Monitor: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Monitor_Info) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Monitor_Fab) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tipo de monitor: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Monitor_Type) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next

MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Multimídia ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Multimídia</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>DISPOSITIVO DE ÁUDIO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, ProductName, Manufacturer"
objClass = "Win32_SoundDevice"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colSound = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colSound
Sound_Name = objItem.Name
Sound_ProdName = objItem.ProductName
Sound_Manufacturer = objItem.Manufacturer
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nome do dispositivo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Sound_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nome do Produto: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Sound_ProdName) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Sound_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Sound_Name = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>DISPOSITIVO DE MÍDIA</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, Drive, Manufacturer, Description"
objClass = "Win32_CDROMDrive"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colCD = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colCD
CD_Name = objItem.Name
CD_Drive = objItem.Drive
CD_Manufacturer = objItem.Manufacturer
CD_Description = objItem.Description
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nome do dispositivo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CD_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Letra do dispositivo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CD_Drive) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CD_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Descrição: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(CD_Description) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If CD_Name = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Dispositivos de entrada ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Dispositivos de entrada</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>TECLADO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLfLf

strProperties = "Description, IsLocked, Status"
objClass = "Win32_Keyboard"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colKeyb = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colKeyb
Keyb_Dev = objItem.Description
Keyb_Lock = objItem.IsLocked
Keyb_Status = objItem.Status
If Keyb_Lock = True Then
Keyb_Lock = "Bloqueado"
Else
Keyb_Lock = "Desbloqueado"
End If
If Keyb_Status = "OK" Then
Keyb_Status = "Ativo"
Else
Keyb_Status = "Desativado"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Teclado: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Keyb_Dev) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Bloqueio: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Keyb_Lock) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Status: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Keyb_Status) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>MOUSE</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLfLf

strProperties = "HardwareType, DeviceInterface, Manufacturer, NumberOfButtons, Handedness"
objClass = "Win32_PointingDevice"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colMouse = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colMouse
Mouse_Dev = objItem.HardwareType
Mouse_IntConn = objItem.DeviceInterface
Mouse_Manufacturer = objItem.Manufacturer
Mouse_NumButtons = objItem.NumberOfButtons
Mouse_Type = objItem.PointingType
Mouse_Resolution = objItem.Handedness
If Mouse_IntConn = 1 Then
Mouse_IntConn = "Outra"
ElseIf Mouse_IntConn = 2 Then
Mouse_IntConn = "Não definida"
ElseIf Mouse_IntConn = 3 Then
Mouse_IntConn = "Serial"
ElseIf Mouse_IntConn = 4 Then
Mouse_IntConn = "PS/2"
ElseIf Mouse_IntConn = 5 Then
Mouse_IntConn = "Infra Vermelho"
ElseIf Mouse_IntConn = 6 Then
Mouse_IntConn = "HP-HIL"
ElseIf Mouse_IntConn = 7 Then
Mouse_IntConn = "Bus Mouse"
ElseIf Mouse_IntConn = 8 Then
Mouse_IntConn = "ADB (Apple Desktop Bus)"
ElseIf Mouse_IntConn = 160 Then
Mouse_IntConn = "DB-9"
ElseIf Mouse_IntConn = 161 Then
Mouse_IntConn = "micro-DIN"
ElseIf Mouse_IntConn = 162 Then
Mouse_IntConn = "USB"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Mouse: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Mouse_Dev) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Interface de conexão: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Mouse_IntConn) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Fabricante: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Mouse_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Número de botões: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Mouse_NumButtons) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Resolução: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Mouse_Resolution) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Informações de Storage ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Storage</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>UNIDADES FÍSICAS</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Model, InterfaceType, Partitions, Size, Status"
objClass = "Win32_DiskDrive"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colStorage = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colStorage
HD_Name = objItem.Model
HD_Intface = objItem.InterfaceType
HD_Part = objItem.Partitions
HD_Size = FormatValue(objItem.Size)
HD_SMART = objItem.Status
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Disco/Modelo: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(HD_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Interface: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(HD_Intface) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Número de partições: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(HD_Part) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tamanho: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(HD_Size) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>S.M.A.R.T.: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(HD_SMART) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>UNIDADES LÓGICAS</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 84px;'>UNIDADE </span><span class='s_line' style='width: 105px;'>| TIPO </span><span class='s_line' style='width: 105px;'>| FORMATAÇÃO </span><span class='s_line' style='width: 125px;'>| TOTAL </span><span class='s_line' style='width: 125px;'>| LIVRE</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, FileSystem, Size, FreeSpace, DriveType"
objClass = "Win32_LogicalDisk"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colDisks = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colDisks
Name_Storage = objItem.Name
File_System = objItem.FileSystem
Total_Space = FormatValue(objItem.Size)
Free_Space = FormatValue(objItem.FreeSpace)
Select Case objItem.DriveType
Case 0: Disk_Type = "Não encontrado"
Case 1: Disk_Type = "RAM Disk"
Case 2: Disk_Type = "Removível"
Case 3: Disk_Type = "Fixo"
Case 4: Disk_Type = "Drive de Rede"
Case 5: Disk_Type = "CD-ROM/DVD-ROM"
End Select
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 84px;'>" & CheckNull(Name_Storage) & " </span><span class='li_itens' style='width: 105px;'>| " & CheckNull(Disk_Type) & " </span><span class='li_itens' style='width: 105px;'>| " & CheckNull(File_System) & " </span><span class='li_itens' style='width: 125px;'>| " & CheckNull(Total_Space) & " </span><span class='li_subitens' style='width: 125px;'>| " & CheckNull(Free_Space) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Informações de Tape Drives ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Tape/Dat Drives</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, Manufacturer, Compression, ECC"
objClass = "Win32_TapeDrive"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colTape = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in ColTape
Tape_Name = objItem.Name
Tape_Manufacturer = objItem.Manufacturer
Tape_Compression = objItem.Compression
Tape_ECC = objItem.ECC
If Tape_Compression = True Then
Tape_Compression = "Ativada"
Else
Tape_Compression = "Desativada"
End If
If Tape_ECC = True Then
Tape_ECC = "Suportado"
Else
Tape_ECC = "Não suportado"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Tape Drive: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Tape_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Fabricante: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Tape_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Compressão: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Tape_Compression) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Checagem de erro por hardware: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Tape_ECC) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Tape_Name = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Impressoras ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Impressora</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 195px;'>ÍTEM</span><span class='s_line' style='width: 349px;'>DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "DriverName, ShareName, HorizontalResolution, VerticalResolution, PrinterState"
objClass = "Win32_Printer"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colInstalledPrinters = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colInstalledPrinters 
Print_DRVName = objItem.DriverName
Print_ShareName = objItem.ShareName
Print_HResol = objItem.HorizontalResolution
Print_VResol = objItem.VerticalResolution
Print_State = objItem.PrinterState
If Print_State = 0 Then
Print_State = "Sem impressão"
ElseIf Print_State = 1 Then
Print_State = "Pausada"
ElseIf Print_State = 2 Then
Print_State = "Erro"
ElseIf Print_State = 3 Then
Print_State = "Pendente de deleção"
ElseIf Print_State = 4 Then
Print_State = "Paper Jam"
ElseIf Print_State = 5 Then
Print_State = "Saída de papel"
ElseIf Print_State = 6 Then
Print_State = "Manual Feed"
ElseIf Print_State = 7 Then
Print_State = "Problema com papel"
ElseIf Print_State = 8 Then
Print_State = "Offline"
ElseIf Print_State = 9 Then
Print_State = "IO Ativo"
ElseIf Print_State = 10 Then
Print_State = "Busy"
ElseIf Print_State = 11 Then
Print_State = "Imprimindo"
ElseIf Print_State = 12 Then
Print_State = "Output Bin Full"
ElseIf Print_State = 13 Then
Print_State = "Não avaliável"
ElseIf Print_State = 14 Then
Print_State = "Aguardando"
ElseIf Print_State = 15 Then
Print_State = "Processando"
ElseIf Print_State = 16 Then
Print_State = "Inicialização"
ElseIf Print_State = 17 Then
Print_State = "Atenção - Perigo"
ElseIf Print_State = 18 Then
Print_State = "Cartucho baixo"
ElseIf Print_State = 19 Then
Print_State = "Nenhum cartucho"
ElseIf Print_State = 20 Then
Print_State = "Page Punt"
ElseIf Print_State = 21 Then
Print_State = "Intervenção do usuário é requerida"
ElseIf Print_State = 22 Then
Print_State = "Sem memória"
ElseIf Print_State = 23 Then
Print_State = "Tampa aberta"
ElseIf Print_State = 24 Then
Print_State = "Servidor não disponível neste momento"
ElseIf Print_State = 25 Then
Print_State = "Economia de Energia"
Else
Print_State = "-"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Impressora: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Print_DRVName) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nome do compartilhamento: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Print_ShareName) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Resolução horizontal: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Print_HResol) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Resolução vertical: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Print_VResol) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Estado da impressora: </span><span class='li_itens' style='width: 349px;'>" & CheckNull(Print_State) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Print_DRVName = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Dispositivos de rede ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Conectividade</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>REDE</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLfLf

strProperties = "ProductName, Manufacturer, MACAddress, AdapterType"
objClass = "Win32_NetworkAdapter"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE AdapterType='Ethernet 802.3' AND ProductName !='Packet Scheduler Miniport'"
Set colNet = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colNet
Net_Prod = objItem.ProductName
Net_Manufacturer = objItem.Manufacturer
Net_MAC = objItem.MACAddress
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Dispositivo de LAN: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Net_Prod) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Fabricante: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Net_Manufacturer) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Endereço MAC: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Net_MAC) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Net_Prod = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "<li><span class='s_line' style='width: 544px;'>MODEM</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Caption, DeviceType, AttachedTo, DriverDate, StatusInfo"
objClass = "Win32_PotsModem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colModem = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colModem
Modem_Name = objItem.Caption
Modem_Type = objItem.DeviceType
Modem_Port = objItem.AttachedTo
Modem_DRVDate = FormatDataTime(objItem.DriverDate)
Modem_Status = objItem.StatusInfo
If Modem_Status = 1 Then
Modem_Status = "Outro"
ElseIf Modem_Status = 2 Then
Modem_Status = "Não identificável"
ElseIf Modem_Status = 3 Then
Modem_Status = "Ativo"
ElseIf Modem_Status = 4 Then
Modem_Status = "Desativado"
ElseIf Modem_Status = 5 Then
Modem_Status = "Não aplicável"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Modem: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Modem_Name) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Tipo: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Modem_Type) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Porta: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Modem_Port) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Data do Driver: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Modem_DRVDate) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 180px;'>Status do disposotivo: </span><span class='li_itens' style='width: 364px;'>" & CheckNull(Modem_Status) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Modem_Name = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Grupos e Usuários ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Grupos e usuários do sistema</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 160px;'>GRUPOS </span><span class='s_line' style='width: 384px;'>| USUÁRIOS</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

Set colGroups = GetObject("WinNT://" & strPC & "")
colGroups.Filter = Array("group")
For Each objItem In colGroups
Group_Name = UCase(objItem.Name)
Last_Group = Group_Name

Set Grupos = GetObject("WinNT://" & strPC & "/"& Group_Name &", group")
For Each objUser in Grupos.members
User_Name = " ¬> " & objUser.Name
If Last_Group = Group_Name then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 160px;'>" & Group_Name & "</span><span class='li_subitens' style='width: 384px;'>" & User_Name & "</span></li>"
Else
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 160px;'>" & Last_Group & "</span><span class='li_subitens' style='width: 384px;'>" & User_Name & "</span></li>"
End if
User_Name = ""
Last_Group = ""
Next
If Last_Group <> "" then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 160px;'>" & Group_Name & "</span><span class='li_subitens' style='width: 384px;'>" & CheckNull(User_Name) & "</span></li>"
End if
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Group_Name = "" And User_Name = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Compartilhamentos ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Compartilhamentos</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 195px;'>ÍTEM</span><span class='s_line' style='width: 349px;'>DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Path, Name, Description"
objClass = "Win32_Share"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colShares = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colShares
Shares_Path = objItem.Path
Shares_Folder = objItem.Name
Shares_Description = objItem.Description
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Caminho: </span><span class='li_subitens' style='width: 349px;'>" & CheckNull(Shares_Path) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Nome: </span><span class='li_subitens' style='width: 349px;'>" & CheckNull(Shares_Folder) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 195px;'>Descrição: </span><span class='li_subitens' style='width: 349px;'>" & CheckNull(Shares_Description) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Shares_Path = "" AND Shares_Folder = "" AND Shares_Description = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Drives Mapeados ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Drives mapeados</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 160px;'>ÍTEM </span><span class='s_line' style='width: 384px;'>| DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, Description"
objClass = "Win32_LogicalDisk"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE Description = 'Network Connection'"
Set colMaps = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colMaps
Map_Drive = objItem.Name
Map_Description = objItem.Description
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 160px;'>" & CheckNull(Map_Drive) & " </span><span class='li_subitens' style='width: 384px;'>| " & CheckNull(Map_Description) & "</span></li>"
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If Map_Drive = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Softwares instalados ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Softwares</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 260px;'>SOFTWARE </span><span class='s_line' style='width: 100px;'>| VERSÃO</span><span class='s_line' style='width: 184px;'>| CAMINHO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKey, arrSubKeys
For Each strSubKey In arrSubKeys
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKey & "\" & strSubKey, "DisplayName", strDisplayName
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKey & "\" & strSubKey, "DisplayVersion", strDisplayVersion
objReg.GetStringValue HKEY_LOCAL_MACHINE, strKey & "\" & strSubKey, "InstallLocation", strInstallLocation
Soft_Install = strDisplayName
Soft_Version = strDisplayVersion
Soft_Vendor = strInstallLocation
If Soft_Install <> "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 260px;'>" & CheckNull(Soft_Install) & " </span><span class='li_itens' style='width: 100px;'>| " & CheckNull(Soft_Version) & " </span><span class='li_itens' style='width: 184px;'>| " & CheckNull(Soft_Vendor) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
Next
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Hot-Fix e Patches ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Hot-Fix e Patch</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 160px;'>ÍTEM </span><span class='s_line' style='width: 384px;'>| DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "HotFixID, Description"
objClass = "Win32_QuickFixEngineering"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colQuickFixes = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colQuickFixes
Hot_Fix = objItem.HotFixID
Hot_Fix_Description = objItem.Description
If Hot_Fix = "File 1" Then
Else
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 160px;'>" & CheckNull(Hot_Fix) & " </span><span class='li_itens' style='width: 384px;'>| " & CheckNull(Hot_Fix_Description) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
Next
If Hot_Fix = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Variáveis do Sistema ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Variáveis do sistema</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 160px;'>VARIÁVEIS DO SISTEMA </span><span class='s_line' style='width: 384px;'>| DESCRIÇÃO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "Name, VariableValue"
objClass = "Win32_Environment"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colVarSys = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colVarSys
Var_Sys_Name = objItem.Name
Var_Sys_Description = objItem.VariableValue
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 160px;'>" & CheckNull(Var_Sys_Name) & " </span><span class='li_itens' style='width: 384px;'>| " & CheckNull(Var_Sys_Description) & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf

':::::::::::::::::::::::::::::::::::::
':: ::
':: Serviços do Sistema ::
':: ::
':::::::::::::::::::::::::::::::::::::
MountHtml = MountHtml & "<br/>" & vbCrLf
MountHtml = MountHtml & "<div id='div_t1'>" & vbCrLf
MountHtml = MountHtml & "<h3>Serviços do sistema</h3>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
MountHtml = MountHtml & "<ul class='list_itens'>" & vbCrLf
MountHtml = MountHtml & "<li><span class='s_line' style='width: 300px;'>SERVIÇO </span><span class='s_line' style='width: 122px;'>| STATUS</span><span class='s_line' style='width: 122px;'>MODO</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf

strProperties = "DisplayName, State, StartMode"
objClass = "Win32_Service"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colServices = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colServices 
SRV_Name = objItem.DisplayName
SRV_Status = objItem.State
SRV_Status_Mode = objItem.StartMode 
If SRV_Status = "Stopped" Then
SRV_Status = "<span class='red'>Parado</span>"
End If
If SRV_Status = "Running" Then
SRV_Status = "<span class='g reen'>Iniciado</span>"
End If
If SRV_Status_Mode = "Manual" Then
SRV_Status_Mode = "Manual"
ElseIf SRV_Status_Mode = "Disabled" Then
SRV_Status_Mode = "Desabilitado"
ElseIf SRV_Status_Mode = "Auto" Then
SRV_Status_Mode = "Automático"
End If
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 300px;'>" & SRV_Name & " </span><span class='li_itens' style='width: 122px;'>| " & SRV_Status & " </span><span class='li_itens' style='width: 122px;'>| " & SRV_Status_Mode & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
Next
If SRV_Name = "" Then
MountHtml = MountHtml & "<li><span class='li_itens' style='width: 544px;'>" & VarMsg_Err3 & "</span></li>" & vbCrLf
MountHtml = MountHtml & "<span class='l_demarc'></span>" & vbCrLf
End If
MountHtml = MountHtml & "</ul>" & vbCrLf
MountHtml = MountHtml & "<span class='l_foot'></span>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf
End Sub

':::::::::::::::::::::::::::::::::::::
':: ::
':: Sub Fecha Browser ::
':: ::
':::::::::::::::::::::::::::::::::::::
Sub DataG_CloseHTML(byval strOpc)
MountHtml = MountHtml & "</br>" & vbCrLf
MountHtml = MountHtml & "<div id='l_foot'>Término do inventário: " & AudData() & "</div>" & vbCrLf
MountHtml = MountHtml & "<div id='t_foot' align='center'>" & vbCrLf
MountHtml = MountHtml & "<input type='button' value='About' onclick=" & Chr(34) & "confirm('\n" & Var_Soft_Name & Var_Soft_Version & Var_Soft_Build & VarMsg_About &"')" & Chr(34) & "/>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf
MountHtml = MountHtml & "</div>" & vbCrLf
MountHtml = MountHtml

If strOpc = 3 then
objIE.Document.Body.InnerHTML = "<div id='t_foot' align='center'>Arquivo processado com sucesso!</div>"
Else
objIE.Document.Body.InnerHTML = MountHtml
End If
Wrt_HTML.Write "</body>" & vbCrLf
Wrt_HTML.Write "</html>" & vbCrLf
Wrt_HTML.Close
Set Wrt_HTML = Nothing
End Sub

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função inicia Browser ::
':: ::
':::::::::::::::::::::::::::::::::::::
Sub Init()
isValidParameters = False
sComputer = ""
sOpt = ""
sVisible = ""

Set oArgs = WScript.Arguments
For x = 0 To oArgs.Count - 1
nPosicao = Instr(1, oArgs(x), ":", 1)
If (Trim(Ucase(Left(oArgs(x), nPosicao))) = "/PC:") Then sComputer = Mid(oArgs(x), nPosicao + 1, Len(oArgs(x)))
If (Trim(sComputer) = "." Or IsEmpty(Trim(sComputer))) Then
On Error Resume next
Set oShell = WScript.CreateObject("WScript.Shell")
' Recupera o nome do PC no Win9x
sComputer = oShell.RegRead("HKLM\System\CurrentControlSet\Services\VxD\VNETSUP\ComputerName")
' Recupera o nome do PC no WinXP ou Win2000
sComputer = oShell.RegRead("HKLM\System\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")
Set oShell = Nothing
On Error goto 0
End if
If (Trim(Ucase(Left(oArgs(x), nPosicao))) = "/OPT:") Then sOpt = Mid(oArgs(x), nPosicao + 1, Len(oArgs(x)))
If (Trim(Ucase(Left(oArgs(x), nPosicao))) = "/REL:") Then sVisible = Mid(oArgs(x), nPosicao + 1, Len(oArgs(x)))
If (Trim(Ucase(Left(oArgs(x), nPosicao))) = "/DXF:") Then sVisible = Mid(oArgs(x), nPosicao + 1, Len(oArgs(x)))
Next
If sComputer <> "" And sOpt <> "" And sVisible <> "" Then isValidParameters = True
Start_Verify_Pc isValidParameters, sComputer, sOpt, sVisible
End Sub

Function Start_Verify_Pc(byval haveParameters, byval strPC, byval strOpc, byval strHTML)
If Not haveParameters Then
Do
strPC = InputBox(Var_Msg_1, Var_Soft_Name)
If strPC = "0" Then wScript.Quit(1)
Loop Until strPC <> ""
Flag = False
Do
strOpc = InputBox(Var_Msg_2, Var_Soft_Name)
If strOpc = "0" Then wScript.Quit(1)
If IsNumeric(strOpc) Then
strOpc = Cint(strOpc)
If ((strOpc >= 1) And (strOpc <= 3)) Then
Flag = True
End If
End if
Loop Until Flag = True

If strHTML = "1" Then
strHTML = True
ElseIf strHTML = "0" Then
strHTML = False
ElseIf strHTML = "" Then
strHTML = True
End If
End If
StartIE strHTML
DG_PCInfo strPC
DataG_CloseHTML strOpc
If strOpc > 1 Then
DataG_CreateFileHTML strPc
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função Cria Cabeçalho HMTL ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function HeaderHtml()
Header_Html = Header_Html & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & "?>" & vbCrLf
Header_Html = Header_Html & "<!DOCTYPE html PUBLIC " & Chr(34) & "-//W3C//DTD XHTML 1.1 Strict//EN" & Chr(34) & " " & Chr(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd" & Chr(34) & ">" & vbCrLf
Header_Html = Header_Html & "<html xmlns=" & Chr(34) & "http://www.w3.org/1999/xhtml" & Chr(34) & " xml:lang=" & Chr(34) & "pt-br" & Chr(34) & " lang=" & Chr(34) & "pt-br" & Chr(34) & ">" & vbCrLf
Header_Html = Header_Html & "<head>" & vbCrLf
Header_Html = Header_Html & "<title>" & Var_Soft_Name & Var_Soft_Version & Var_Soft_Build & "</title>" & vbCrLf
Header_Html = Header_Html & "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & "/>" & vbCrLf
Header_Html = Header_Html & "</head>" & vbCrLf
Header_Html = Header_Html & "<style type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf
Header_Html = Header_Html & "<!--" & vbCrLf
Header_Html = Header_Html & vbCrLf
Header_Html = Header_Html & "body" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:11px;" & vbCrLf
Header_Html = Header_Html & " background-color:#1e77d3;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff',endColorStr='#1e77d3',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "#t_head, #t_foot" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:19px;" & vbCrLf
Header_Html = Header_Html & " background-color:#cadef4;" & vbCrLf
Header_Html = Header_Html & " color:black;" & vbCrLf
Header_Html = Header_Html & " position:static;" & vbCrLf
Header_Html = Header_Html & " top:0px; left:0px;" & vbCrLf
Header_Html = Header_Html & " width:650px;" & vbCrLf
Header_Html = Header_Html & " height:50px;" & vbCrLf
Header_Html = Header_Html & " padding:10 0 0 0;" & vbCrLf
Header_Html = Header_Html & " border:1px;" & vbCrLf
Header_Html = Header_Html & " border-left:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-right:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-top:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-bottom:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " text-align: center;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff',endColorStr='#cadef4',gradientType='0');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "#l_head" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:11px;" & vbCrLf
Header_Html = Header_Html & " background-color:#cadef4;" & vbCrLf
Header_Html = Header_Html & " color:black;" & vbCrLf
Header_Html = Header_Html & " position:static;" & vbCrLf
Header_Html = Header_Html & " top:0px; left:0px;" & vbCrLf
Header_Html = Header_Html & " width:650px;" & vbCrLf
Header_Html = Header_Html & " height:15px;" & vbCrLf
Header_Html = Header_Html & " padding:0 0 0 2;" & vbCrLf
Header_Html = Header_Html & " border-left:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-right:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-bottom:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " text-align: left;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff',endColorStr='#cadef4',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "#l_foot" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:11px;" & vbCrLf
Header_Html = Header_Html & " background-color:#cadef4;" & vbCrLf
Header_Html = Header_Html & " color:black;" & vbCrLf
Header_Html = Header_Html & " position:static;" & vbCrLf
Header_Html = Header_Html & " top:0px; left:0px;" & vbCrLf
Header_Html = Header_Html & " width:650px;" & vbCrLf
Header_Html = Header_Html & " height:15px;" & vbCrLf
Header_Html = Header_Html & " padding:0 0 0 2;" & vbCrLf
Header_Html = Header_Html & " border-left:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-right:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-top:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " text-align: left;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff',endColorStr='#cadef4',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "#t_body" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family: Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size: 11px;" & vbCrLf
Header_Html = Header_Html & " background-color: #f5f5f5;" & vbCrLf
Header_Html = Header_Html & " color: black;" & vbCrLf
Header_Html = Header_Html & " position: static;" & vbCrLf
Header_Html = Header_Html & " width: 650px;" & vbCrLf
Header_Html = Header_Html & " padding:0 0 0 0;" & vbCrLf
Header_Html = Header_Html & " text-align: center;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#cadef4',endColorStr='#ffffff',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "#div_t1" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:11px;" & vbCrLf
Header_Html = Header_Html & " background-color:#f0f0f0;" & vbCrLf
Header_Html = Header_Html & " width:550px;" & vbCrLf
Header_Html = Header_Html & " padding: 2 2 2 2;" & vbCrLf
Header_Html = Header_Html & " margin: 0;" & vbCrLf
Header_Html = Header_Html & " border: 1px;" & vbCrLf
Header_Html = Header_Html & " border-right:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-top:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-left:gray 1px solid;" & vbCrLf
Header_Html = Header_Html & " border-bottom:gray 1px solid;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff',endColorStr='#f5f5f5',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "h3" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " text-align: left;" & vbCrLf
Header_Html = Header_Html & " background-color:#1e77d3;" & vbCrLf
Header_Html = Header_Html & " padding: 7 7 7 7;" & vbCrLf
Header_Html = Header_Html & " color:white;" & vbCrLf
Header_Html = Header_Html & " position:static;" & vbCrLf
Header_Html = Header_Html & " height:40px;" & vbCrLf
Header_Html = Header_Html & " margin-bottom: 0px;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#1e77d3',endColorStr='#ffffff',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "ul" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " margin: 0px;" & vbCrLf
Header_Html = Header_Html & " background-color:#ffffff;" & vbCrLf
Header_Html = Header_Html & " text-align: left;" & vbCrLf
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & ".l_demarc" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:0px;" & vbCrLf
Header_Html = Header_Html & " background-color:#1e77d3;" & vbCrLf
Header_Html = Header_Html & " position: relative;" & vbCrLf
Header_Html = Header_Html & " width: 100%;" & vbCrLf
Header_Html = Header_Html & " margin-top: 2px;" & vbCrLf
Header_Html = Header_Html & " margin-bottom: 2px;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#1e77d3',endColorStr='#ffffff',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & "li.list_itens" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " display: inline;" & vbCrLf
Header_Html = Header_Html & " list-style-type: none;" & vbCrLf
Header_Html = Header_Html & " text-align: left;" & vbCrLf
' Header_Html = Header_Html & " margin: 0px;" & vbCrLf
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & ".li_itens" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " vertical-align: middle;" & vbCrLf
Header_Html = Header_Html & " padding: 2 4 2 4;" & vbCrLf
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & ".s_line" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " font-size:10px;" & vbCrLf
Header_Html = Header_Html & " color:#ffffff;" & vbCrLf
Header_Html = Header_Html & " background-color:green;" & vbCrLf
Header_Html = Header_Html & " height: 20px;" & vbCrLf
Header_Html = Header_Html & " margin: 0px;" & vbCrLf
Header_Html = Header_Html & " padding: 2 0 0 5;" & vbCrLf
Header_Html = Header_Html & " position: static;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='green',endColorStr='#ffffff',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & ".red" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " color:red;" & vbCrLf
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & ".green" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " color:green;" & vbCrLf
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & ".l_foot" & vbCrLf
Header_Html = Header_Html & " {" & vbCrLf
Header_Html = Header_Html & " font-family:Trebuchet MS;" & vbCrLf
Header_Html = Header_Html & " text-align: center;" & vbCrLf
Header_Html = Header_Html & " background-color:#1e77d3;" & vbCrLf
Header_Html = Header_Html & " color:white;" & vbCrLf
Header_Html = Header_Html & " position:static;" & vbCrLf
Header_Html = Header_Html & " width: 100%;" & vbCrLf
Header_Html = Header_Html & " height: 20px;" & vbCrLf
Header_Html = Header_Html & " padding: 5 5 5 5;" & vbCrLf
Header_Html = Header_Html & " text-align: left;" & vbCrLf
If strDx = True Then
Header_Html = Header_Html & " filter: progid:DXImageTransform.Microsoft.Gradient(startColorStr='#1e77d3',endColorStr='#ffffff',gradientType='1');" & vbCrLf
End If
Header_Html = Header_Html & " }" & vbCrLf
Header_Html = Header_Html & vbCrLf
Header_Html = Header_Html & "-->" & vbCrLf
Header_Html = Header_Html & "</style>" & vbCrLf
Header_Html = Header_Html & "<body>" & vbCrLf

HeaderHtml = Header_Html
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função Formatar Valor Memória ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatValue(objFormatMem)
If objFormatMem <> 0 then 
If CDbl(objFormatMem) < 1024^3 Then 
If CDbl(objFormatMem) < 1024^2 Then 
Mem_Divisor = 1024
Mem_Unit = " KB" 
Else
Mem_Divisor = 1024^2 
Mem_Unit = " MB" 
End If 
Else 
Mem_Divisor = 1024^3 
Mem_Unit = " GB" 
End If
If Mem_Divisor = 1024 Then
FormatValue = FormatNumber(objFormatMem / Mem_Divisor, 0) & Mem_Unit
ElseIf Mem_Divisor = 1024^2 Then
FormatValue = FormatNumber(objFormatMem / Mem_Divisor, 0) & Mem_Unit
Else
FormatValue = FormatNumber(objFormatMem / Mem_Divisor, 1) & Mem_Unit
End If
Else
FormatValue = "-" 
End If 
End Function

Function MemValue(VarComplCheck)
If VarComplCheck <> 0 Then 
If VarComplCheck < 1024 Then
MemValue = Clng(VarComplCheck) & " KB"
ElseIf VarComplCheck > 1023 Then
MemValue = Clng(VarComplCheck /1024) & " MB"
End If
Else
VarComplCheck = "-"
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função verifica valores nulos ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function CheckNull(VarForCheck)
If IsNull(VarForCheck) = True Or VarForCheck = "" Or VarForCheck = " " Then
CheckNull = "-"
Else
CheckNull = VarForCheck
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função formatar data ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatDataTime(VarDateCheck)
LeftStr = Left(VarDateCheck, 8)
DYear = Left(LeftStr, 4)
DMonth = Mid(LeftStr, 5, 2)
DDay = Right(LeftStr, 2)
FormatDataTime = DDay & "/" & DMonth & "/" & DYear
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função formatar voltagem ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatVolt(VarVoltCheck)
If VarVoltCheck <> 0 Then
FormatVolt = Replace(FormatNumber(VarVoltCheck / 10, 1), ",", ".") & " V"
Else
FormatVolt = "-"
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função formatar Bit ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatBit(VarBitCheck)
If VarBitCheck <> 0 Then
FormatBit = VarBitCheck & " Bit"
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função formatar porcentagem ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatPerc(VarPercCheck)
If VarPercCheck => 0 Then
FormatPerc = VarPercCheck & " %"
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função formatar Hertz ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatHz(VarHzCheck)
If VarHzCheck => 0 Then
FormatHz = VarHzCheck & " Hz"
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função formatar Clock ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function FormatClock(VarClockCheck)
If VarClockCheck > 1 Then
FormatClock = VarClockCheck & " MHz"
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função Data Auditoria ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function AudData()
Date_Time = Date & " às " & Time
AudData = Date_Time
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função de Tratamento de Erro ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function Trat_Err()
If Err.Number <> 0 Then
MsgBox VarMsg_Err2, 16, Err.Number & vbCrLf & "-" & vbCrLf & Err.Description
ObjIE.Quit
wScript.Quit
Err.Clear
End If
End Function

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função Cria Relatório HTML ::
':: ::
':::::::::::::::::::::::::::::::::::::
Sub DataG_CreateFileHTML(byval strPc)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strPC & "\root\cimv2")
strProperties = "Name"
Set colPC = objWMIService.ExecQuery("SELECT " & strProperties & " FROM Win32_ComputerSystem", , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colPC
Host_name = objItem.Name
Date_Time = Replace(Date, "/", "-") & "~" & Replace(Time,":", "-")
htmlRel = UCase(Host_name & "-" & Date_Time & ".html")

Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objFH = objFileSys.CreateTextFile(htmlRel)
objFH.WriteLine HeaderHtml()
objFH.WriteLine MountHtml
objFH.WriteLine "</body>" & vbCrLf
objFH.WriteLine "</html>"
objFH.close
Set oFileSys = Nothing
Set objFH = Nothing
wScript.Quit
Next
End Sub

':::::::::::::::::::::::::::::::::::::
Init() ::
':::::::::::::::::::::::::::::::::::::
