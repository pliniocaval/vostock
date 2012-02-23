strComputer = "LocalHost"
 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objDhcpNic = objWMIService.ExecQuery _
("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
 
For Each objNic in objDhcpNic
objNic.ReleaseDHCPLease()
Next
 
For Each objNic in objDhcpNic
objNic.RenewDHCPLease()
Next