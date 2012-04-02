'Script Para Geração de Inventario | Leonardo Vivas
' --------------------------------------------------

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oSysInfo = CreateObject("ADSystemInfo")

'Captura e volta 1 nivel do diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")
For i = 0 to Ubound(arrPath) - 1
    DIRS = DIRS & arrPath(i) & "\"
Next 
oShell.CurrentDirectory = DIRS

'msgbox "Não parar em caso de erros"
On Error Resume Next

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
  
ChecaArquivoSai(USERLOGS & "\Inventario-" & COMP & ".log")

sDN = oSysInfo.DomainDNSName
sUserDN = oSysInfo.UserName
Set objUser = GetObject("LDAP://" & sDN & "/" & sUserDN)

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & COMP & "\root\cimv2")
  
Set colComputer = objWMIService.ExecQuery _
("Select * from Win32_ComputerSystem")

Set colComputerIP = objWMIService.ExecQuery _
("Select * from Win32_NetworkAdapterConfiguration")
 
Set colSystemInfo = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem",,48)

strQuery = "SELECT * FROM Win32_Printer"
Set colInstalledPrinters = objWMIService.ExecQuery _
(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)

strProperties = "Model, InterfaceType, Partitions, Size, Status"
objClass = "Win32_DiskDrive"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colStorage = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)

strProperties = "Product, Manufacturer, Model, OtherIdentifyingInfo, SerialNumber, PartNumber, Version"
objClass = "Win32_BaseBoard"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colMBoard = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)

strProperties = "Name, Manufacturer, BuildNumber, CurrentLanguage, ReleaseDate, SerialNumber, SMBIOSBIOSVersion, Version"
objClass = "Win32_BIOS"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colBios = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)


For Each objComputer in colComputer
strUserName = "Usuário Logado: " & objComputer.UserName
strHostName = "Estação: " & objComputer.Name
PC_Type = "Tipo do sistema: " & objComputer.SystemType
PC_Mem = "Memória do sistema: " & FormatValue(objComputer.TotalPhysicalMemory)
PC_DOMAIN = "Dominio: " & objComputer.Domain
PC_FABRIC = "Fabricante: " & objComputer.Manufacturer
PC_MODEL = "Modelo: " & objComputer.Model
Next

strDisplayName = "Nome do usuário: " & trim(objUser.DisplayName)
If strDisplayName = "" Then
strDisplayName = strUserName
End If

For Each objItem in colBios
Bios_Name = "Bios : " & objItem.Name
Bios_Manufacturer = "Fabricante da Bios : " & objItem.Manufacturer
Bios_Build = objItem.BuildNumber
Bios_Lang = objItem.CurrentLanguage
Bios_ReleaseDate = "Data da Bios: " & FormatDataTime(objItem.ReleaseDate)
Bios_SN = "Serial da Bios : "& objItem.SerialNumber
Bios_SMBiosVersion = objItem.SMBIOSBIOSVersion
Bios_Version = objItem.Version
Next
 
For Each IPConfig in colComputerIP
If Not IsNull(IPConfig.IPAddress) Then
For intIPCount = LBound(IPConfig.IPAddress) _
to UBound(IPConfig.IPAddress)
strIPAddress = strIPAddress & "End. de IP: " & IPConfig.IPAddress(intIPCount) & "~" & vbCrLf
next
If InStr(strMACAddress, "MAC Address: " & IPConfig.MACAddress & "~") = 0 Then
strMACAddress = strMACAddress & "MAC Address: " & IPConfig.MACAddress & "~" & vbCrLf
End If
End If

Next
 
If Right(strIPAddress, 1) = "~" Then
strIPAddress = Left(strIPAddress, Len(strIPAddress) - 1)
End If
strIPAddress = Replace(strIPAddress, "~", " ")
strMACAddress = Replace(strMACAddress, "~", " ")
For Each objItem in colSystemInfo
strOS_Caption = "S.O.: " & objItem.Caption
strOS_SPVersion = "Service Pack: " & objItem.CSDVersion
strOS_VerNumber = "Versão do S.O.: " & objItem.Version
SO_Serial = "Número serial: " & objItem.SerialNumber
Next

For Each objItem in colInstalledPrinters 
Print_DRVName = Print_DRVName & "Impressora Padrão: " & objItem.DriverName & vbCrLf

next

For Each objItem in colStorage
HD_Name = "Disco/Modelo: " & objItem.Model
HD_Intface = "Interface: " & objItem.InterfaceType
HD_Part = "Número de partições: " & objItem.Partitions
HD_Size = "Tamanho: " & FormatValue(objItem.Size)
HD_SMART = "S.M.A.R.T.: " & objItem.Status
next

For Each objItem in colMBoard
MB_Product = objItem.Product
MB_Manufacturer = objItem.Manufacturer
MB_Model = objItem.Model
MB_NS = "Numero de Serie: " & objItem.SerialNumber
MB_PartNumber = "Part Number: " & objItem.PartNumber
MB_Version = objItem.Version
MB_OtherInfo = objItem.OtherIdentifyingInfo
Next

LOGONSERV = "Logon Server: " & LOGON

strLogFile = USERLOGS & "\Inventario-" & COMP & ".log"
arrTipos = split(arrTipos,";")
Set strLogFile = oFso.OpenTextFile(strLogFile, 8, True, 0)
strLogFile.WriteLine  VBCRLF
strLogFile.WriteLine "==================================================="
strLogFile.WriteLine "Iniciando do Inventario EM: " & now
strLogFile.WriteLine "==================================================="  
strLogFile.WriteLine  VBCRLF
strLogFile.WriteLine CheckNull(LOGONSERV)
strLogFile.WriteLine CheckNull(PC_DOMAIN)
strLogFile.WriteLine "Estação: " & CheckNull(COMP)
strLogFile.WriteLine CheckNull(strUserName)
strLogFile.WriteLine CheckNull(strDisplayName)
'strLogFile.WriteLine CheckNull(strHostStatus)
strLogFile.WriteLine CheckNull(strIPAddress)
strLogFile.WriteLine CheckNull(strMACAddress)
strLogFile.WriteLine CheckNull(strOS_Caption)
strLogFile.WriteLine CheckNull(PC_Type)
strLogFile.WriteLine CheckNull(strOS_SPVersion)
strLogFile.WriteLine CheckNull(strOS_VerNumber)
strLogFile.WriteLine CheckNull(SO_Serial)
strLogFile.WriteLine CheckNull(PC_FABRIC)
strLogFile.WriteLine CheckNull(PC_MODEL)
strLogFile.WriteLine CheckNull(MB_PartNumber)
strLogFile.WriteLine CheckNull(MB_NS)
strLogFile.WriteLine CheckNull(Bios_Manufacturer)
strLogFile.WriteLine CheckNull(Bios_SN)
strLogFile.WriteLine CheckNull(Bios_ReleaseDate)
strLogFile.WriteLine CheckNull(PC_Mem)
strLogFile.WriteLine CheckNull(HD_Name)
strLogFile.WriteLine CheckNull(HD_Intface)
strLogFile.WriteLine CheckNull(HD_Part)
strLogFile.WriteLine CheckNull(HD_Size)
strLogFile.WriteLine CheckNull(HD_SMART)
strLogFile.WriteLine CheckNull(Print_DRVName)
strLogFile.WriteLine "==================================================="
strLogFile.WriteLine "Fim do Inventario EM: " & now
strLogFile.WriteLine "==================================================="  

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

':::::::::::::::::::::::::::::::::::::
':: ::
':: Função verifica valores nulos ::
':: ::
':::::::::::::::::::::::::::::::::::::
Function CheckNull(VarForCheck)
If IsNull(VarForCheck) = True Or VarForCheck = "" Or VarForCheck = " " Then
CheckNull = " N/A "
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