'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 21/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")

' Não parar em caso de erros
On Error Resume Next

'Carregando variaveis
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes
' SAIDA POR Exceção
If left(ucase(computador),4)="PVIL" Then wscript.quit
If left(ucase(computador),3)="MXM" Then
	If left(ucase(user),4)="PVIL" Then
	ass
	Else
	wscript.quit
	End if
Else
ass
End If


'--------------------
Function ass
' Não parar em caso de erros
On Error Resume Next
'Realiza limpeza das assinaturas antigas
set folder = objFSO.getFolder (vAPPDATA &"\Microsoft\Signatures\")   
for each file in folder.files
File.delete
next
set folder = objFSO.getFolder (vAPPDATA &"\Microsoft\Assinaturas\")   
for each file in folder.files
File.delete
next
objfso.deletefolder vAPPDATA & "\Microsoft\Signatures\*.*",true
objfso.deletefolder vAPPDATA & "\Microsoft\Assinaturas\*.*",true

'msgbox "Carregando variaveis (variaveis contidas em outro arquivo)"
' Caso não queira usar um arquivo externo apague ou comente as linhas 19 à 24.
' Remova o comentario da linha 25,  edite a linha 26 e coloque a pasta na rede onde estão as imagens, leia o comentario da linha 33, verifique as linhas 98 e 99 estas tambem são sobre a imagem.
varfile = "\\cemusadobrasil.com.br\SYSVOL\cemusadobrasil.com.br\SCRIPTS\Logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes = f.ReadAll
  f.close
  execute constantes
'sUserDN = objSysInfo.UserName
'TISRV = \\csrv06\ti$
'fim do carregamento de variaveis

Set objUser = GetObject("LDAP://" & sUserDN)
arrDept = split(sUserDN, ",")

'Informação da UO
sLocation = mid(arrDept(2), 4) 'identifica a localização do usuario baseado na UO. Mude o valor "4" para definir a profundidade desta. o nome da UO sera o Nome da imagem.
'Informação do AD
strName = objUser.FullName
strTitle = objUser.Description
strPhone = objUser.TelephoneNumber
strCel = objUser.mobile
strDepartment = objUser.Department 

'Criação da Assinatura
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objRange = objDoc.Range()
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

' Criando email

'Nome
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(0, 0, 0)
If strName <> "" Then
objSelection.TypeText "   " & strName
objSelection.TypeText(Chr(11))
Else
objSelection.TypeText "   ERRO AO GERAR ASSINATURA DE EMAIL FAVOR CONTACTAR O SUPORTE"
objSelection.TypeText(Chr(11))
End If

'Cargo
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(0, 0, 0)
If strTitle <> "" Then
objSelection.TypeText "   " & strTitle
objSelection.TypeText(Chr(11))
'Else
'objSelection.TypeText "   ERRO AO GERAR ASSINATURA DE EMAIL FAVOR CONTACTAR O SUPORTE"
'objSelection.TypeText(Chr(11))
End If

'Departamento
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(0, 0, 0)
If strDepartment <> "" Then
objSelection.TypeText "   Departamento de " & strDepartment
objSelection.TypeText(Chr(11))
End If

'Telefone Fixo
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(38, 38, 38)
If strPhone <> "" Then
objSelection.TypeText "   Tel.:" & strPhone
objSelection.TypeText(Chr(11))
End If

'Telefone Movel
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(38, 38, 38)
If strCel <> "" Then
objSelection.TypeText "   Cel.:" & strCel
objSelection.TypeText(Chr(11))
End If

objSelection.TypeText(Chr(11))

'Imagem (busca a imagem na rede)
If objFSO.FileExists(TISRV & "\ASS\" & sLocation & ".jpg") Then 'verifica de existe a imagem que ira aparecer na assinatura.
Set s = objSelection.InlineShapes.AddPicture(TISRV & "\ASS\" & sLocation & ".jpg")	'aqui voce define o caminho da imagem que ira aparecer na assinatura.
With s

End With
objSelection.TypeParagraph()
Else
'caso não exista imagem insira os dados que que ser exibido em caso de falha no carregamento da imagem.
Set s = objSelection.InlineShapes.AddPicture(TISRV & "\ASS\Generica.jpg")	'aqui voce define o caminho da imagem que ira aparecer na assinatura.
With s

End With
objSelection.TypeParagraph()
End If

'messagem de confidencialidade(BR)
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(0, 0, 0)
objSelection.TypeText "Aviso de confidencialidade"
objSelection.TypeText(Chr(11))
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.italic = True
objSelection.Font.Size = 7
objSelection.Font.Color = RGB(128, 128, 128)
objSelection.TypeText "Esta mensagem, incluindo seus eventuais anexos, pode conter informações confidenciais, de uso restrito e/ou legalmente protegidas. Se você recebeu esta mensagem por engano, não deve usar, copiar, divulgar, distribuir ou tomar qualquer atitude com base nestas informações. Solicitamos que você elimine a mensagem imediatamente de seu sistema e avise ao remetente respondendo a mensagem.  Todas as opiniões, conclusões ou informações contidas nesta mensagem somente serão consideradas como provenientes da CEMUSA ou de suas subsidiárias quando efetivamente confirmadas, formalmente, por um de seus representantes legais, devidamente autorizados para tanto."
objSelection.TypeText(Chr(11))


'Menagem sobre Meio Ambiente(BR)
objSelection.Font.Name = "Webdings"
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(35, 142, 35)
objSelection.Font.Bold = False
objSelection.Font.italic = False
objSelection.TypeText "P "
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 8
objSelection.TypeText "Antes de imprimir este email pense se é realmente necessario."
objSelection.TypeParagraph()

'messagem de confidencialidade(US)
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(0, 0, 0)
objSelection.TypeText "Confidentiality Note"
objSelection.TypeText(Chr(11))
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.italic = true
objSelection.Font.Size = 7
objSelection.Font.Color = RGB(128, 128, 128)
objSelection.TypeText "Privileged/Confidential Information may be contained in this message. If you are not the addressee indicated in this message (or responsible for delivery of the message to such person), you may not copy or deliver this message to anyone. In such case, you should destroy this message and kindly notify the sender by reply email. Please advise immediately if you or your employer does not consent to email or messages of this kind. Opinions, conclusions and other information in this message that do not relate to the official business of CEMUSA shall be understood as neither given nor endorsed by it."
objSelection.TypeText(Chr(11))


'Menagem sobre Meio Ambiente(US)
objSelection.Font.Name = "Webdings"
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(35,142,35)
objSelection.Font.Bold = False
objSelection.Font.italic = False
objSelection.TypeText "P "
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 8
objSelection.TypeText "Before printing this e-mail, think if it is necessary."

objSignatureEntries.Add "Cemusa", objRange
objSignatureObject.NewMessageSignature = "Cemusa"
objSignatureObject.ReplyMessageSignature = "Cemusa"
objDoc.Saved = True
objWord.Quit
End Function