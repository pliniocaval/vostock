'Script do Assinatura
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

ApagaArquivosPastas(vAPPDATA &"\Microsoft\Signatures\") 
ApagaArquivosPastas(vAPPDATA &"\Microsoft\Assinaturas\")  

Const END_OF_STORY = 6
Const wdFormatHTML = 8

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
objTable.Cell(2, 1).Range.InlineShapes.AddPicture DIR & "\IMG\agente.jpg"
objTable.Cell(2, 1).Range.TypeText(Chr(11))
objTable.Cell(2, 1).Range.InlineShapes.AddPicture DIR & "\IMG\Platinum.jpg"
objTable.Columns(1).Width = objWord.InchesToPoints(1)
.ParagraphFormat.Alignment = wdAlignParagraphRight
	  
objTable.Cell(2, 2).Select

.Font.Name = "Verdana"
.Font.Size = 8
.Font.Bold = True
.Font.Color = RGB(128, 128, 128)

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
.TypeText "   " & "Platinum Investimentos | " & "Assessor de Investimentos"
.TypeText(Chr(11))

'Tel e Site
If strPhone <> "" Then    
.TypeText "   " &  strPhone & " | "
.Hyperlinks.Add .range, strWeb
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

'fb / tw
.TypeText "   "
.InlineShapes.AddPicture DIR & "\IMG\fb.jpg"
.InlineShapes.AddPicture DIR & "\IMG\tw.jpg"
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
objSelection.Font.Bold = False
objSelection.TypeText "atua no mercado financeiro através da XP Investimentos CCTVM S/A, realizando o atendimento de pessoas físicas e jurídicas(não-institucionais). Na forma da legislação da CVM, o agente autônomo de investimento não pode administrar ou gerir o patrimônio de investidores. O agente autônomo é um intermediário e depende da autorização prévia do cliente para realizar operações no mercado financeiro."
objSelection.TypeText(Chr(11))
objSelection.TypeText " Esta mensagem, incluindo os seus anexos, contém informações confidenciais destinadas a indivíduo e propósito específicos, sendo protegida por lei. Caso você não seja a pessoa a quem foi dirigida a mensagem, deve apagá-la. É terminantemente proibida a utilização, acesso, cópia ou divulgação não autorizada das informações presentes nesta mensagem."
objSelection.TypeText(Chr(11))
objSelection.TypeText "As informações contidas nesta mensagem e em seus anexos são de responsabilidade de seu autor, não representando necessariamente ideias, opiniões, pensamentos ou qualquer forma de posicionamento por parte da "
objSelection.Font.Bold = True
objSelection.TypeText "Platinum Investimentos – Agente Autônomo de Investimentos Ltda. "
objSelection.TypeText(Chr(11))
objSelection.Font.Bold = False
objSelection.TypeText "O investimento em ações é um investimento de risco e rentabilidade passada não é garantia de rentabilidade futura. Na realização de operações com derivativos existe a possibilidade de perdas superiores aos valores investidos, podendo resultar em significativas perdas patrimoniais."
objSelection.TypeText(Chr(11))
objSelection.TypeText "Para informações e dúvidas, favor contatar seu operador."
objSelection.TypeParagraph()
objSelection.TypeText "Para reclamações, favor contatar a Ouvidoria da XP Investimentos no telefone nº 0800-722-3710."
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

Set objSelection = objDoc.Range()
objSignatureEntries.Add "Ass Padrao", objSelection
objSignatureObject.NewMessageSignature = "Ass Padrao"
objSignatureObject.ReplyMessageSignature = "Ass Padrao"
objDoc.Saved = True
objWord.Quit