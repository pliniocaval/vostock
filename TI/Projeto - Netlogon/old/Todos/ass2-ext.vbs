'Criação de Assinatura Padão para o outlook
'versão 1.0
'03/06/2006

On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")
Set objnet = CreateObject("WScript.Network")

strUser = objSysInfo.UserName
'Wscript.Echo strUser
Set objUser = GetObject("LDAP://" & strUser)
With objUser
  strName = .FullName
  strTitle = .Description
End With

strlogon = objNet.UserName
strCompany = objUser.Company
strAddress = objUser.streetAddress
strpostalCode = objUser.postalCode
strl = objUser.l
strco = objUser.co
strPhone = objUser.TelephoneNumber
strFax = objUser.facsimileTelephoneNumber
strMail = objuser.mail
strWeb = objuser.wWWHomePage
strCel = objUser.mobile

Set objword = CreateObject("Word.Application")
With objword

  Set objDoc = .Documents.Add()
  Set objSelection = .Selection
  Set objEmailOptions = .EmailOptions
End With

Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
With objSelection

  .ParagraphFormat.Alignment = wdAlignParagraphRight
  .TypeParagraph

  With .Font
    .Name = "Arial"
    .Size = 10
    .Bold = false
  End With
    .TypeText strName & Chr(11)
  With .Font
    .Name = "Arial"
    .Size = 10
    .Bold = False
    .Italic = False
  End With
    .TypeText strTitle & Chr(11)
    
    objSelection.Font.Size = "10" 
    objSelection.Font.Name = "Arial"    
    objSelection.Font.Bold = False    
    .TypeText strCompany & Chr(11)

  With .Font
    .Name = "Arial"
    .Size = 10
    .Bold = false
  End With
    .Font.Italic = False
    .TypeText strAddress & Chr(11) & strl & " - " & strco & Chr(11) & "CEP:" & strpostalCode & Chr(11) & "Tel.:" & strPhone & Chr(11) & "Fax.:" & strFax & Chr(11) & "Email:" & strMail & Chr(11) & "Site: " & strWeb & Chr(11)
    .TypeText Chr(11)
    .InlineShapes.AddPicture "\\cemusadobrasil.com.br\geral\email.bmp", True, True
	objSelection.Font.Size = "10"
    objSelection.Font.italic = False
    objSelection.Font.Color = 4612846
    objSelection.Font.Bold = True 
    .TypeText Chr(11)
    .TypeText Chr(11)
objSelection.Font.Size = "8"
    objSelection.Font.italic = true
    objSelection.Font.Color = 8421504
    objSelection.Font.Bold = False  
    objSelection.TypeText "Esta mensagem, incluindo seus eventuais anexos, pode conter informações confidenciais, de uso restrito e/ou legalmente protegidas. Se você recebeu esta mensagem por engano, não deve usar, copiar, divulgar, distribuir ou tomar qualquer atitude com base nestas informações. Solicitamos que você elimine a mensagem imediatamente de seu sistema e avise ao remetente respondendo a mensagem.  Todas as opiniões, conclusões ou informações contidas nesta mensagem somente serão consideradas como provenientes da CEMUSA ou de suas subsidiárias quando efetivamente confirmadas, formalmente, por um de seus representantes legais, devidamente autorizados para tanto."

    .TypeText Chr(11)
    .TypeText Chr(11)
objSelection.Font.Size = "8"
    objSelection.Font.italic = true
    objSelection.Font.Color = 8421504    
    objSelection.Font.Bold = False  
    objSelection.TypeText "Privileged/Confidential Information may be contained in this message. If you are not the addressee indicated in this message (or responsible for delivery of the message to such person), you may not copy or deliver this message to anyone. In such case, you should destroy this message and kindly notify the sender by reply email. Please advise immediately if you or your employer does not consent to email or messages of this kind. Opinions, conclusions and other information in this message that do not relate to the official business of CEMUSA shall be understood as neither given nor endorsed by it. "

	.TypeText Chr(11)
    .TypeText Chr(11)
objSelection.Font.Size = "8"
    objSelection.Font.italic = true
    objSelection.Font.Color = 8421504    
    objSelection.Font.Bold = False  
    objSelection.TypeText "Este e-mail y cualquiera de sus ficheros anexos son confidenciales y pueden incluir información privilegiada. Si usted no es el destinatario adecuado (o responsable de remitirlo a la persona indicada), agradeceríamos lo notificase/reenviase inmediatamente al emisor. No revele estos contenidos a ninguna otra persona, no los utilice para otra finalidad, ni almacene y/o copie esta información en medio alguno. Opiniones, conclusiones y otro tipo de información relacionada con este mensaje que no sean relativas a la actividad propia de CEMUSA, deberán ser entendidas exclusivas del emisor."

  End With

Set objSelection = objDoc.Range()
objSignatureEntries.Add "Cemusa-ext", objSelection
objSignatureObject.NewMessageSignature = "Cemusa-int"
objSignatureObject.ReplyMessageSignature = "Cemusa-int"
objDoc.Saved = True
objword.Quit
objword.Quit
wscript.quit