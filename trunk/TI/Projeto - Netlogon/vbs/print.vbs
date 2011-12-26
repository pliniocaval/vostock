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

' Não parar em caso de erros
On Error Resume Next

'msgbox "Carregando variaveis"
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

Set objUser = GetObject("WinNT://" & Domain & "/" & user)
' Adicionar impressoras
For each oGroup in objUser.Groups
    If UCase(oGroup.Name) = "COPACABANA" Then
	    'wscript.echo  "Impressoras (Adicionar)"
	       if not Mapeada("\\csrv01\HP4010") then objnet.AddWindowsPrinterConnection "\\csrv01\HP4010"
		   if not Mapeada("\\CRJ022\HP930C") then objnet.AddWindowsPrinterConnection "\\CRJ022\HP930C"
	    'wscript.echo "Impressoras (Remover)"
		   if Mapeada("\\csrv01\Canon iR2016 UFRII LT") then objnet.RemovePrinterConnection "\\csrv01\Canon iR2016 UFRII LT"
	End If
	If UCase(oGroup.Name) = "SÃO CRISTOVÃO" Then
	    'wscript.echo  "Impressoras (Adicionar)"
	      if not Mapeada("\\sqlscpi\HP4014") then objnet.AddWindowsPrinterConnection "\\sqlscpi\HP4014"
		  if not Mapeada("\\sqlscpi\P3005") then objnet.AddWindowsPrinterConnection "\\sqlscpi\P3005"
		'wscript.echo "Impressoras (Remover)"
		if Mapeada("\\CEMUSA007\HP4014") then objnet.RemovePrinterConnection "\\CEMUSA007\HP4014"
		if Mapeada("\\sqlscpi\HPColor2") then objnet.RemovePrinterConnection "\\sqlscpi\HPColor2"
	End If
		
' Definir impressora Padrão
If UCase(oGroup.Name) = "FINANCEIRO" Then
		'wscript.echo "Impressora (Padrão)"
	       if Mapeada("\\csrv01\HP4010") then objnet.SetDefaultPrinter "\\csrv01\HP4010"
	End If
If UCase(oGroup.Name) = "COMPRAS" Then
		'wscript.echo "Impressora (Padrão)"
	       if Mapeada("\\sqlscpi\P3005") then objnet.SetDefaultPrinter "\\sqlscpi\P3005"
	End If
If UCase(oGroup.Name) = "SCPI" Then
		'wscript.echo "Impressora (Padrão)"
	       if Mapeada("\\sqlscpi\P3005") then objnet.SetDefaultPrinter "\\sqlscpi\P3005"
	End If
If UCase(oGroup.Name) = "PRODUCAO" Then
		'wscript.echo "Impressora (Padrão)"
	       if Mapeada("\\sqlscpi\HP4014") then objnet.SetDefaultPrinter "\\sqlscpi\HP4014"
	End If
Next


'funçoes
Function Mapeada(Caminho)
 Mapeada = False
 Set objNet = WScript.CreateObject("WScript.Network")
 Set colPrinters = objNet.EnumPrinterConnections
 For i = 0 to colPrinters.Count -1 Step 2
 if ucase(colPrinters.Item (i + 1)) = ucase(caminho) then
 Mapeada = True
 exit for
 end if
 Next
end function