'Criação de Assinatura Padão para o outlook
'versão 1.0
'03/06/2006

On Error Resume Next
Set objnet = CreateObject("WScript.Network")
Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
'Wscript.Echo strUser
Set objUser = GetObject("LDAP://" & strUser)
With objUser
  strName = .FullName
  strTitle = .Description
End With


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
    .TypeText "Tel.:" & strPhone & Chr(11)
    .TypeText Chr(11)

  End With

Set objSelection = objDoc.Range()
objSignatureEntries.Add "Cemusa-int", objSelection
objSignatureObject.NewMessageSignature = "Cemusa-int"
objSignatureObject.ReplyMessageSignature = "Cemusa-int"
objDoc.Saved = True
objword.Quit
objword.Quit
wscript.quit