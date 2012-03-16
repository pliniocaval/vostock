'Script Para Geração de Relatorio tipos de arquivos especificos
'autoria Leonardo Vivas
'Versão 2.0
'criação 03/06/2009
'modificação 03/03/2012
' -----------------------------------------------------------------' 

Set oNet = CreateObject("WScript.Network")
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Captura e volta 1 nivel do diretorio
DIRE = oFSO.GetParentFolderName(WScript.ScriptFullName)
arrPath = Split(DIRE, "\")

For i = 0 to Ubound(arrPath) - 1
    DIR = DIR & arrPath(i) & "\"
Next 

oShell.CurrentDirectory = DIR

'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = DIR & "\SYS\LOGON.INI"
  Set SYS = oFSO.OpenTextFile(varfile)
  SYSFILE =   SYS.ReadAll
  SYS.close
  execute SYSFILE

'msgbox "Carregando arquivo de Funções"
varfile = DIR & "\SYS\FNC.INI"
  Set FNC = oFSO.OpenTextFile(varfile)
  FNCFILE =   FNC.ReadAll
  FNC.close
  execute FNCFILE
  
'msgbox "Carregando arquivo de Parametrização"
varfile = DIR & "\SYS\EMP.INI"
  Set EMP = oFSO.OpenTextFile(varfile)
  EMPFILE =   EMP.ReadAll
  EMP.close
  execute EMPFILE
  
ChecaArquivoSai(USERLOGS & "\Arquivos-" & COMP & ".txt")

Dim intTotalSpace, intTotalSpacemp3, intTotalSpaceavi, intTotalSpacewmv, intTotalSpacempeg
Dim numeroCont, numeroContmp3, numeroContavi, numeroContwmv, numeroContmpeg, numeroContPST
Dim intFileSizemp3, intFileSizeavi, intFileSizewmv, intFileSizempeg, intFileSizePST

intTotalSpace=0
intTotalSpacemp3=0
intTotalSpaceavi=0
intTotalSpacewmv=0
intTotalSpacempeg=0
intTotalSpacePST=0

numeroCont=0
numeroContmp3=0
numeroContavi=0
numeroContwmv=0
numeroContmpeg=0
numeroContPST=0

intFileSizemp3=0
intFileSizeavi=0
intFileSizewmv=0
intFileSizempeg=0
intFileSizePST=0
  
set objTextFile = oFso.CreateTextFile(USERLOGS & "\Arquivos-" & COMP & ".txt", True)

objTextFile.writeline "Estação: " & COMP
objTextFile.writeline "Usuario atual: " & user
Set objWMIService = GetObject("winmgmts:\\").ExecQuery("SELECT * FROM CIM_DataFile WHERE Drive = 'C:'")

For Each objItem in objWMIService
select case objItem.Extension
case "mp3"
numeroContmp3 = numeroContmp3 +1
intFileSizemp3=objItem.FileSize
intTotalSpacemp3= intTotalSpacemp3 + intFileSizemp3
call escrever()
'CopyArquivo objItem.Name
case "avi"
numeroContavi = numeroContavi +1
intFileSizeavi=objItem.FileSize
intTotalSpaceavi= intTotalSpaceavi + intFileSizeavi
call escrever()
'CopyArquivo objItem.Name
case "wmv"
numeroContwmv = numeroContwmv +1
intFileSizewmv=objItem.FileSize
intTotalSpacewmv= intTotalSpacewmv + intFileSizewmv
call escrever()
'CopyArquivo objItem.Name
case "mpeg"
numeroContmpeg = numeroContmpeg +1
intFileSizempeg=objItem.FileSize
intTotalSpacempeg= intTotalSpacempeg + intFileSizempeg
call escrever()
'CopyArquivo objItem.Name
case "PST"
numeroContPST = numeroContPST +1
intFileSizePST=objItem.FileSize
intTotalSpacePST= intTotalSpacePST + intFileSizePST
call escrever()
'CopyArquivo objItem.Name
end select 

Next
objTextFile.writeline "Total de Arquivos MP3: " & CheckNull(numeroContmp3) & " " &"Tamanho: " & FormatValue(intTotalSpacemp3)
objTextFile.writeline "Total de Arquivos AVI: " & CheckNull(numeroContavi) & " " & "Tamanho: " & FormatValue(intTotalSpaceavi)
objTextFile.writeline "Total de Arquivos WMV: " & CheckNull(numeroContwmv) & " " &"Tamanho: " & FormatValue(intTotalSpacewmv)
objTextFile.writeline "Total de Arquivos MPEG: " & CheckNull(numeroContmpeg) & " " &"Tamanho: " & FormatValue(intTotalSpacempeg)
objTextFile.writeline "Total de Arquivos PST: " & CheckNull(numeroContPST) & " " &"Tamanho: " & FormatValue(intTotalSpacePST)
objTextFile.writeline "Total de Arquivos: " & CheckNull(numeroCont)
objTextFile.writeline "Total Ocupado: " & FormatValue(intTotalSpace)
'Wscript.echo "Pronto" 

function escrever()
Set objetoSF2 = oFso.GetFile (objItem.name)
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
Set oFso = CreateObject("Scripting.FileSystemObject")
Set objFile = oFso.GetFile(arquivo)
Set objnet = CreateObject("WScript.Network")


'****** AINDA EM TESTE******
'MsgBox "DIRETORIO DESTINO DOS ARQUIVOS"
dirdest = "c:\TESTE\"
If Not oFso.FolderExists(dirdest) Then oFso.CreateFolder(dirdest)
pastapai = RIGHT(objFile.ParentFolder,LEN(objFile.ParentFolder)-3)
arrTipos = split(pastapai,"\")
For x = 0 to UBOUND(arrTipos)
   if oFso.folderexists(dirdest & arrTipos(x)) = false Then
      oFso.CreateFolder(dirdest & arrTipos(x))
      dirDest = dirDest & arrTipos(x) & "\"
   Else
      dirDest = dirDest & arrTipos(x) & "\"
   End if
Next
oFso.CopyFile objFile.Path  , dirdest
End sub