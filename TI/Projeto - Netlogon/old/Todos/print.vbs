'Script do logon
'autoria Leonardo Vivas
'Vers�o 1.8
'cria��o 03/06/2009
'modifica��o 14/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' N�o parar em caso de erros
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
	    'wscript.echo "Impressoras (Remover)"
		   if Mapeada("\\csrv01\Canon iR2016 UFRII LT") then objnet.RemovePrinterConnection "\\csrv01\Canon iR2016 UFRII LT"
	End If
' Definir impressora Padr�o
If UCase(oGroup.Name) = "CONTABILIDADE" Then
		'wscript.echo "Impressora (Padr�o)"
	       if Mapeada("\\csrv01\HP4010") then objnet.SetDefaultPrinter "\\csrv01\HP4010"
	End If

Next


'fun�oes
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