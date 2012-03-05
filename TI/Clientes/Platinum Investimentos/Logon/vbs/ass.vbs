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
varfile = "C:\Users\lvivas\Desktop\Vostock-Projects\TI\Clientes\Platinum Investimentos\Logon\Logon.ini"
Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set f = objFSO.OpenTextFile(varfile)
  constantes =   f.ReadAll
  f.close
  execute constantes

' SAIDA POR Exceção
if left(ucase(computador),4)="SRV" then wscript.quit
ass

'--------------------
Function ass
' Não parar em caso de erros
On Error Resume Next
'Realiza limpeza das assinaturas antigas
set folder = objFSO.getFolder (vAPPDATA &"\Microsoft\Signatures\")   
for each file in folder.files
'File.delete
next
set folder = objFSO.getFolder (vAPPDATA &"\Microsoft\Assinaturas\")   
for each file in folder.files
'File.delete
next
'objfso.deletefolder vAPPDATA & "\Microsoft\Signatures\*.*",true
'objfso.deletefolder vAPPDATA & "\Microsoft\Assinaturas\*.*",true

Set objUser = GetObject("LDAP://" & sUserDN)

'Informação da UO
arrDept = split(sUserDN, ",")
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
objDoc.Tables.Add objRange,1,2
Set objTable = objDoc.Tables(1)
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries



' Criando email
With objSelection

objTable.Rows.Add()

'Imagem (busca a imagem na rede)
If objFSO.FileExists(TISRV & "\ASS\image002.jpg") Then 'verifica de existe a imagem que ira aparecer na assinatura.
Set s = objSelection.InlineShapes.AddPicture(TISRV & "\ASS\image002.jpg")	'aqui voce define o caminho da imagem que ira aparecer na assinatura.
With s
.Height = 80
.Width = 80
End With
objSelection.TypeParagraph()
Else
'caso não exista imagem insira os dados que que ser exibido em caso de falha no carregamento da imagem.
Set s = objSelection.InlineShapes.AddPicture(TISRV & "\ASS\Generica.jpg")	'aqui voce define o caminho da imagem que ira aparecer na assinatura.
With s
End With
.ParagraphFormat.Alignment = wdAlignParagraphRight
objSelection.TypeParagraph()
End If

objTable.Columns(1).Width = objWord.InchesToPoints(1)

'Nome
objTable.Cell(1, 2).Range.Font.Name = "Verdana"
objTable.Cell(1, 2).Range.Font.Bold = True
objTable.Cell(1, 2).Range.Font.Size = 9
objTable.Cell(1, 2).Range.Font.Color = RGB(0, 0, 0)
If strName <> "" Then
objTable.Cell(1, 2).Range.TypeText "   " & strName
objTable.Cell(1, 2).Range.TypeText(Chr(11))
Else
objTable.Cell(1, 2).Range.TypeText "   ERRO AO GERAR ASSINATURA DE EMAIL FAVOR CONTACTAR O SUPORTE"
objTable.Cell(1, 2).Range.TypeText(Chr(11))
End If

'Cargo
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.Size = 9
objSelection.Font.Color = RGB(0, 0, 0)
If strTitle <> "" Then
objSelection.TypeText "   " & strTitle
objSelection.TypeText(Chr(11))
Else
objSelection.TypeText "   ERRO AO GERAR ASSINATURA DE EMAIL FAVOR CONTACTAR O SUPORTE"
objSelection.TypeText(Chr(11))
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

objTable.Columns(2).Width = objWord.inchesToPoints(0) 
objSelection.EndKey END_OF_STORY
objTable.close

'messagem de confidencialidade(BR)
objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(0, 0, 0)
objSelection.TypeText "A Platinum Investimentos – Agente Autônomo de Investimentos Ltda."

objSelection.Font.Name = "Verdana"
objSelection.Font.Bold = False
objSelection.Font.italic = True
objSelection.Font.Size = 8
objSelection.Font.Color = RGB(128, 128, 128)
objSelection.TypeText "é uma empresa de agentes autônomos de investimento devidamente registrada na Comissão de Valores Mobiliários e credenciada, na forma da Instrução Normativa n. 497/11. A relação completa dos sócios agentes autônomos da"
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
End With

objSignatureEntries.Add "Platinum", objRange
objSignatureObject.NewMessageSignature = "Platinum"
objSignatureObject.ReplyMessageSignature = "Platinum"
objDoc.Saved = True
objWord.Quit
End Function