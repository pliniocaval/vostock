Const END_OF_STORY = 6
Const wdFormatHTML = 8

On Error Resume Next
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
With objUser
  strName = .FullName
  strTitle = .Description
End With

strCompany = objUser.Company
strl = objUser.l
strco = objUser.co
strPhone = objUser.TelephoneNumber
strFax = objUser.facsimileTelephoneNumber
strMobile = objUser.Mobile
strWeb = objuser.wWWHomePage
strUserName = objuser.sAMAccountName

Set objword = CreateObject("Word.Application")
With objword

  Set objDoc = .Documents.Add()
  Set objSelection = .Selection
  Set objEmailOptions = .EmailOptions
  
  Set objRange = objDoc.Range()
  objDoc.Tables.Add objRange,1,2
  Set objTable = objDoc.Tables(1)

End With

Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

With objSelection

objTable.Rows.Add()

objTable.Cell(2, 1).Range.InlineShapes.AddPicture "F:\Vostock - Projects\Clientes\Platinum\Netlogon\VBS\agente.jpg"
.TypeText(Chr(11))
objTable.Cell(2, 1).Range.InlineShapes.AddPicture "F:\Vostock - Projects\Clientes\Platinum\Netlogon\VBS\Platinum.jpg"
objTable.Columns(1).Width = objWord.InchesToPoints(1)
.ParagraphFormat.Alignment = wdAlignParagraphRight
	  
objTable.Cell(2, 2).Select
With .Font
.Name = "Verdana"
.Size = 8
.Bold = True
.Color = RGB(128, 128, 128)
End With

'Arrumando
.TypeParagraph()

'Nome
If strName <> "" Then
.TypeText "   " & strName
.TypeText(Chr(11))
Else
.TypeText "   ERRO AO GERAR ASSINATURA"
.TypeText(Chr(11))
End If

'Cargo
If strTitle <> "" Then    
.TypeText "   " & "Platinum Investimentos | " & strTitle
.TypeText(Chr(11))
Else
.TypeText "   " & "Platinum Investimentos | " & "ERRO AO GERAR ASSINATURA"
.TypeText(Chr(11))
End If

'Tel e Site
If strPhone <> "" Then    
.TypeText "   " &  strPhone & " | " & strWeb
.TypeText(Chr(11))
Else
.TypeText "   " & "ERRO AO GERAR ASSINATURA" & " | " 
.Hyperlinks.Add .range, "www.platinuminvest.com.br"
.TypeParagraph()
End If

'Endereço
.TypeText "   " &  "Rio de Janeiro | RJ:" & Chr(11)
With .Font
.Name = "Verdana"
.Size = 8
.Bold = False
.Color = RGB(128, 128, 128)
End With

.TypeText "   " &  "Matriz – Filial – Città América Office" & Chr(11)
.TypeText "   " &  "Av. das Américas, 700 | 2º Andar | Sala 201" & Chr(11)
.TypeText "   " &  "Barra da Tijuca – 22430-041" & Chr(11)

'Arrumando
.TypeParagraph()
.ParagraphFormat.Alignment = wdAlignParagraphRight    
objTable.Columns(2).Width = objWord.inchesToPoints(0) 

objSelection.EndKey END_OF_STORY   

objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 7
objSelection.Font.Color = RGB(128, 128, 128)
objSelection.Font.Bold = False
objSelection.TypeText "A "
objSelection.Font.Bold = True
objSelection.TypeText "Platinum Investimentos – Agente Autônomo de Investimentos Ltda. "
objSelection.Font.Bold = False
objSelection.TypeText "é uma empresa de agentes autônomos de investimento devidamente registrada na Comissão de Valores Mobiliários e credenciada, na forma da Instrução Normativa n. 497/11. A relação completa dos sócios agentes autônomos da "
objSelection.Font.Bold = True
objSelection.TypeText "Platinum Investimentos – Agente Autônomo de Investimentos Ltda. "
objSelection.Font.Bold = False
objSelection.TypeText ", bem como dos demais agentes autônomos contratados pela XP Investimentos Corretora pode ser consultada no site "
objSelection.Hyperlinks.Add .range, "www.cvm.gov.br"
objSelection.TypeText " > > Agentes Autônomos > Relação dos Agentes Autônomos contratados por uma Instituição Financeira > Corretoras > XP Investimentos ou diretamente no site da XP Investimentos CCTVM S/A através do link << "
objSelection.Hyperlinks.Add .range, "www.xpi.com.br"
objSelection.TypeText " >>>. A "
objSelection.Font.Bold = True
objSelection.TypeText "Platinum Investimentos – Agente Autônomo de Investimentos Ltda. "
objSelection.Font.Bold = True
objSelection.TypeText(Chr(11))

'Menagem sobre Meio Ambiente(BR)
objSelection.Font.Name = "Webdings"
objSelection.Font.Size = 7
objSelection.Font.Color = RGB(35, 142, 35)
objSelection.Font.Bold = False
objSelection.Font.italic = False
objSelection.TypeText "P "
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 7
objSelection.TypeText "Antes de imprimir este email pense se é realmente necessario."
objSelection.TypeParagraph()

End With



Set objSelection = objDoc.Range()
objSignatureEntries.Add "Padrao Leonardi", objSelection
objSignatureObject.NewMessageSignature = "Padrao Leonardi"
objSignatureObject.ReplyMessageSignature = "Padrao Leonardi"
objDoc.Saved = True
objWord.Quit