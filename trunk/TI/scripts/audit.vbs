'Script para enumerar Arquivos no sistema local
'Autor: Leonardo vivas
'Versão: 1.0
'Criação: 10/05/2011
'Ultima Modificação: N/A
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

'msgbox "Não parar em caso de erros"
'On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\Logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes = f.ReadAll
  f.close
  execute constantes

Dim intTotalSpace, intTotalSpacemp3, intTotalSpaceavi, intTotalSpacewmv, intTotalSpacempeg
Dim numeroCont, numeroContmp3, numeroContavi, numeroContwmv, numeroContmpeg
Dim intFileSizemp3, intFileSizeavi, intFileSizewmv, intFileSizempeg

intTotalSpace=0
intTotalSpacemp3=0
intTotalSpaceavi=0
intTotalSpacewmv=0
intTotalSpacempeg=0

numeroCont=0
numeroContmp3=0
numeroContavi=0
numeroContwmv=0
numeroContmpeg=0

intFileSizemp3=0
intFileSizeavi=0
intFileSizewmv=0
intFileSizempeg=0
  
set objTextFile = objFSO.CreateTextFile("Arquivos.txt", True)

objTextFile.writeline "Estação: " & computador
objTextFile.writeline "Usuario atual: " & user
Set objWMIService = GetObject("winmgmts:\\").ExecQuery( _
"SELECT * FROM CIM_DataFile WHERE Drive = 'C:'")
For Each objItem in objWMIService
select case objItem.Extension
case "mp3"
numeroContmp3 = numeroContmp3 +1
intFileSizemp3=objItem.FileSize
intTotalSpacemp3= intTotalSpacemp3 + intFileSizemp3
call escrever()
CopyArquivo objItem.Name
case "avi"
numeroContavi = numeroContavi +1
intFileSizeavi=objItem.FileSize
intTotalSpaceavi= intTotalSpaceavi + intFileSizeavi
call escrever()
CopyArquivo objItem.Name
case "wmv"
numeroContwmv = numeroContwmv +1
intFileSizewmv=objItem.FileSize
intTotalSpacewmv= intTotalSpacewmv + intFileSizewmv
call escrever()
CopyArquivo objItem.Name
case "mpeg"
numeroContmpeg = numeroContavi +1
intFileSizempeg=objItem.FileSize
intTotalSpacempeg= intTotalSpacempeg + intFileSizempeg
call escrever()
CopyArquivo objItem.Name
end select 

Next
objTextFile.writeline "Total de Arquivos MP3: " & CheckNull(numeroContmp3) & " " &"Tamanho: " & FormatValue(intTotalSpacemp3)
objTextFile.writeline "Total de Arquivos AVI: " & CheckNull(numeroContavi) & " " & "Tamanho: " & FormatValue(intTotalSpaceavi)
objTextFile.writeline "Total de Arquivos WMV: " & CheckNull(numeroContwmv) & " " &"Tamanho: " & FormatValue(intTotalSpacewmv)
objTextFile.writeline "Total de Arquivos MPEG: " & CheckNull(numeroContmpeg) & " " &"Tamanho: " & FormatValue(intTotalSpacempeg)
objTextFile.writeline "Total de Arquivos: " & CheckNull(numeroCont)
objTextFile.writeline "Total Ocupado: " & FormatValue(intTotalSpace)
'Wscript.echo "Pronto" 

function escrever()
Set objetoSF2 = objFSO.GetFile (objItem.name)
ArquivoInfo = "Arquivo: " & objItem.name & vbCrLf & "Tamanho :" & FormatValue(objItem.filesize) & vbCrLf & "Date last modified:" & objetoSF2.DateLastModified & vbCrLf
objTextFile.writeline ArquivoInfo
numeroCont = numeroCont +1
intFileSize=objItem.FileSize
intTotalSpace= intTotalSpace + intFileSize
end function 

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
FormatValue = "0" 
End If 
End Function 

Function CheckNull(VarForCheck)
If IsNull(VarForCheck) = True Or VarForCheck = "" Or VarForCheck = " " Then
CheckNull = " 0 "
Else
CheckNull = VarForCheck
End If
End Function

Sub CopyArquivo(arquivo)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(arquivo)
Set objnet = CreateObject("WScript.Network")
'****** AINDA EM TESTE******
'MsgBox "DIRETORIO DESTINO DOS ARQUIVOS"
dirdest = "c:\TESTE\"
If Not objFso.FolderExists(dirdest) Then objFso.CreateFolder(dirdest)
pastapai = RIGHT(objFile.ParentFolder,LEN(objFile.ParentFolder)-3)
arrTipos = split(pastapai,"\")
For x = 0 to UBOUND(arrTipos)
   if objFSO.folderexists(dirdest & arrTipos(x)) = false Then
      objFSO.CreateFolder(dirdest & arrTipos(x))
      dirDest = dirDest & arrTipos(x) & "\"
   Else
      dirDest = dirDest & arrTipos(x) & "\"
   End if
Next
'MsgBox dirDest
'MsgBox objFile.Path
objFSO.CopyFile objFile.Path  , dirdest
End sub