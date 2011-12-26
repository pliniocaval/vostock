' Author Leonardo Vivas
' Versão 0.5 - Maio 2009
' -----------------------------------------------------------------' 
Option Explicit

Dim objFSO, objLogFile, objNetwork, objShell, strText, intAns
Dim intConstants, intTimeout, strTitle, intCount, blnLog
Dim strUserName, strComputerName, strIP, strShare, strLogFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
On Error Resume Next
strUserName = objNetwork.UserName
strComputerName = objNetwork.ComputerName
strIP = GetIPAddresses()

strShare = "\\csrv02\logs$\"
strLogFile = strUserName&"-"&strIP&".log"
intTimeout = 20

' Log date/time, user name, computer name, and IP address.
If (objFSO.FolderExists(strShare) = True) Then
    On Error Resume Next
    Set objLogFile = objFSO.OpenTextFile(strShare & "\" _
        & strLogFile, 8, True, 0)
    If (Err.Number = 0) Then
        ' Make three attempts to write to log file.
        intCount = 1
        blnLog = False
        Do Until intCount = 3
            objLogFile.WriteLine "Logoff ; "  & Now & " ; " _
                & strComputerName & " ; " & strUserName & " ; " & strIP
            If (Err.Number = 0) Then
                intCount = 3
                blnLog = True
            Else
                Err.Clear
                intCount = intCount + 1
                If (Wscript.Version > 5) Then
                    Wscript.Sleep 200
                End If
            End If
        Loop
        On Error GoTo 0
        If (blnLog = False) Then
            strTitle = "Logon Error"
            strText = "Log cannot be written."
            strText = strText & vbCrlf _
                & "Another process may have log file open."
            intConstants = vbOKOnly + vbExclamation
            intAns = objShell.Popup(strText, intTimeout, strTitle, _
                intConstants)
        End If
        objLogFile.Close
    Else
        On Error GoTo 0
        strTitle = "Logon Error"
        strText = "Log cannot be written."
        strText = strText & vbCrLf & "User may not have permissions,"
        strText = strText & vbCrLf & "or log folder may not be shared."
        intConstants = vbOKOnly + vbExclamation
        intAns = objShell.Popup(strText, intTimeout, strTitle, intConstants)
    End If
    Set objLogFile = Nothing
End If

' Clean up and exit.

Set objFSO = Nothing
Set objNetwork = Nothing
Set objShell = Nothing

Wscript.Quit

Function GetIPAddresses()
Dim Loc,WbemServices,Adapters,Adapter
Set loc = CreateObject( "WbemScripting.SWbemLocator" )
Set WbemServices = loc.ConnectServer( ,"root\cimv2" )
Set Adapters=WbemServices.ExecQuery( "Select * FROM" & _
" Win32_NetworkAdapterConfiguration" )
For Each Adapter in Adapters
   If NOT IsNull( Adapter.IPAddress) Then
     if  Left(Adapter.IPAddress(0),1) > 0 Then
        strIP = Adapter.IPAddress(0)
     End if
    
   End If
Next
GetIPAddresses = strIP
End Function