'Criação de Assinatura Padão para o outlook
'versão 2.0
'17/03/2011

Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Description
strPhone = objUser.TelephoneNumber

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
'objSelection.Style = No Spacing
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

'Name of Staff
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(0,0,0)
objSelection.TypeText "   " & strName
objSelection.TypeText(Chr(11))

'Role of Staff
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(0,0,0)
objSelection.TypeText "   " & strTitle
objSelection.TypeText(Chr(11))

'Company Contact details
objSelection.Font.Color = RGB(38,38,38)
objSelection.TypeText "   Tel.:" & strPhone
objSelection.TypeParagraph()

'Company Logo (stored in network share accessed by everyone)
'objSelection.InlineShapes.AddPicture("\\csrv01\Geral\cemusa.jpg")
Set s = objSelection.InlineShapes.AddPicture("\\csrv01\Geral\cemusa.jpg")	
With s
.Height = 90
.Width = 485
End With
objSelection.TypeParagraph()

'message confidentiality(BR)
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(0,0,0)
objSelection.TypeText "Aviso de confidencialidade"
objSelection.TypeText(Chr(11))
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.italic = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(128,128,128)
objSelection.TypeText "Esta mensagem, incluindo seus eventuais anexos, pode conter informações confidenciais, de uso restrito e/ou legalmente protegidas. Se você recebeu esta mensagem por engano, não deve usar, copiar, divulgar, distribuir ou tomar qualquer atitude com base nestas informações. Solicitamos que você elimine a mensagem imediatamente de seu sistema e avise ao remetente respondendo a mensagem.  Todas as opiniões, conclusões ou informações contidas nesta mensagem somente serão consideradas como provenientes da CEMUSA ou de suas subsidiárias quando efetivamente confirmadas, formalmente, por um de seus representantes legais, devidamente autorizados para tanto."
objSelection.TypeText(Chr(11))


'environment message(BR)
objSelection.Font.Name = "Webdings"
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(35,142,35)
objSelection.Font.Bold = False
objSelection.Font.italic = False
objSelection.TypeText "P "
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 8
objSelection.TypeText "Antes de imprimir este email pense se é realmente necessario."
objSelection.TypeParagraph()

'message confidentiality(US)
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(0,0,0)
objSelection.TypeText "Confidentiality Note"
objSelection.TypeText(Chr(11))
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.italic = true
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(128,128,128)
objSelection.TypeText "Privileged/Confidential Information may be contained in this message. If you are not the addressee indicated in this message (or responsible for delivery of the message to such person), you may not copy or deliver this message to anyone. In such case, you should destroy this message and kindly notify the sender by reply email. Please advise immediately if you or your employer does not consent to email or messages of this kind. Opinions, conclusions and other information in this message that do not relate to the official business of CEMUSA shall be understood as neither given nor endorsed by it."
objSelection.TypeText(Chr(11))


'environment message(US)
objSelection.Font.Name = "Webdings"
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(35,142,35)
objSelection.Font.Bold = False
objSelection.Font.italic = False
objSelection.TypeText "P "
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 8
objSelection.TypeText "Before printing this e-mail, think if it is necessary."


Set objSelection = objDoc.Range()

objSignatureEntries.Add "cemusa", objSelection
objSignatureObject.NewMessageSignature = "cemusa"
objSignatureObject.ReplyMessageSignature = "cemusa"

objDoc.Saved = True
objWord.Quit