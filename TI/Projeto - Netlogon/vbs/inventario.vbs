'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 14/12/2011
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
  constantes = f.ReadAll
  f.close
  execute constantes

'set file = objFSO.GetFile(LOGUSER &"\" & computador & ".log")		
'If DateDiff("d", file.DateLastModified, Now) > 5 Then 
inventario
'End If


function inventario
 Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & computador & "\root\cimv2")
  
  
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

 
For Each objComputer in colComputer
strUserName = "Usuário Logado: " & objComputer.UserName
strHostName = "Estação: " & objComputer.Name
PC_Type = "Tipo do sistema: " & objComputer.SystemType
PC_Mem = "Memória do sistema: " & FormatValue(objComputer.TotalPhysicalMemory)
PC_DOMAIN = "Dominio: " & objComputer.Domain
PC_FABRIC = "Fabricante: " & objComputer.Manufacturer
PC_MODEL = "Modelo: " & objComputer.Model
Next
 
For Each IPConfig in colComputerIP
If Not IsNull(IPConfig.IPAddress) Then
For intIPCount = LBound(IPConfig.IPAddress) _
to UBound(IPConfig.IPAddress)
strIPAddress = strIPAddress & "End. de IP: " & IPConfig.IPAddress(intIPCount) & "~"
next
end if
Next
 
If Right(strIPAddress, 1) = "~" Then
strIPAddress = Left(strIPAddress, Len(strIPAddress) - 1)
End If
strIPAddress = Replace(strIPAddress, "~", vbCrLf)
 
For Each objItem in colSystemInfo
strOS_Caption = "S.O.: " & objItem.Caption
strOS_SPVersion = "Service Pack: " & objItem.CSDVersion
strOS_VerNumber = "Versão do S.O.: " & objItem.Version
SO_Serial = "Número serial: " & objItem.SerialNumber
Next

For Each objItem in colInstalledPrinters 
Print_DRVName = Print_DRVName & "Impressora: " & objItem.DriverName & vbCrLf
Print_ShareName = Print_ShareName & "Nome do Compartilhamento: " & objItem.ShareName & VBCRLF
next

For Each objItem in colStorage
HD_Name = "Disco/Modelo: " & objItem.Model
HD_Intface = "Interface: " & objItem.InterfaceType
HD_Part = "Número de partições: " & objItem.Partitions
HD_Size = "Tamanho: " & FormatValue(objItem.Size)
HD_SMART = "S.M.A.R.T.: " & objItem.Status
next

LOGONSERV = "Logon Server: " & LOGON

strLogFile = LOGUSER & "\" & computador & ".log"
arrTipos = split(arrTipos,";")
Set strLogFile = objFSO.OpenTextFile(strLogFile, 8, True, 0)
strLogFile.WriteLine  VBCRLF
strLogFile.WriteLine "==================================================="
strLogFile.WriteLine "Iniciando do Inventario EM: " & now
strLogFile.WriteLine "==================================================="  
strLogFile.WriteLine  VBCRLF
strLogFile.WriteLine LOGONSERV
strLogFile.WriteLine PC_DOMAIN
strLogFile.WriteLine strHostStatus
strLogFile.WriteLine strIPAddress
strLogFile.WriteLine strOS_Caption
strLogFile.WriteLine PC_Type
strLogFile.WriteLine strOS_SPVersion
strLogFile.WriteLine strOS_VerNumber
strLogFile.WriteLine SO_Serial
strLogFile.WriteLine PC_FABRIC
strLogFile.WriteLine PC_MODEL
strLogFile.WriteLine strUserName
strLogFile.WriteLine PC_Mem
strLogFile.WriteLine HD_Name 
strLogFile.WriteLine HD_Intface
strLogFile.WriteLine HD_Part
strLogFile.WriteLine HD_Size
strLogFile.WriteLine HD_SMART
strLogFile.WriteLine Print_DRVName
strLogFile.WriteLine "==================================================="
strLogFile.WriteLine "Fim do Inventario EM: " & now
strLogFile.WriteLine "==================================================="  

end function

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