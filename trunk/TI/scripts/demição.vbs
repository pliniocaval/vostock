'Script do logon
'autoria Leonardo Vivas
'Versão 1.8
'criação 03/06/2009
'modificação 21/12/2011
' -----------------------------------------------------------------' 

Set objnet = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'msgbox "Não parar em caso de erros"
On Error Resume Next

'msgbox "Carregando variaveis"
strUsername = objnet.UserName
robocopy = "robocopy.exe"

'===========================================================================================================================
'=======================EDITE AS VARIVEIS CONTIDAS ABAIXO PARA O CORRETO FUNCIONAMENTO DO SCRIPT============================ 
' Local do BKP
BKPSERVER = "\\csrv06\bkp$"
'Letra da unidade de Rede que ira representar o servidor de BKP
strShare = "Z:"  
'pasta do BKP
strBKPfolder = "\BKP"
If Not objFSO.FolderExists(BKPSERVER & "\" & strUsername) Then objFSO.CreateFolder(BKPSERVER & "\" & strUsername)
BKP = BKPSERVER & "\" & strUsername & strBKPfolder
'===========================================================================================================================
strCopyOptions = " /E /COPY:DAT /R:2 /W:10 /TEE /XF *.exe *.pdb *.sfp *.lnk *.db *.rdp /LOG+:" & strShare & "\Backup.log"

'Adicionar Rotina de ping para verificar se servidor esta online (pendente).

If Not objFSO.FolderExists(BKP) Then objFSO.CreateFolder(BKP)
objNet.RemoveNetworkDrive strShare, True, True
objnet.MapNetworkDrive strShare, BKP

'Cria pasta para o BKP

	If Not objFso.FolderExists(strShare) Then objFso.CreateFolder(strShare)
	If Not objFso.FolderExists(strShare & "\Contatos") Then objFso.CreateFolder(strShare & "\Contatos")
    If Not objFso.FolderExists(strShare & "\Desktop") Then objFso.CreateFolder(strShare & "\Desktop")
    If Not objFso.FolderExists(strShare & "\Documentos") Then objFso.CreateFolder(strShare & "\Documentos")
	If Not objFso.FolderExists(strShare & "\Downloads") Then objFso.CreateFolder(strShare & "\Downloads")
    If Not objFso.FolderExists(strShare & "\Favoritos") Then objFso.CreateFolder(strShare & "\Favoritos")
	If Not objFso.FolderExists(strShare & "\Imagens") Then objFso.CreateFolder(strShare & "\Imagens")
	If Not objFso.FolderExists(strShare & "\Musicas") Then objFso.CreateFolder(strShare & "\Musicas")
	If Not objFso.FolderExists(strShare & "\Videos") Then objFso.CreateFolder(strShare & "\Videos")
	If Not objFso.FolderExists(strShare & "\Jogos Salvos") Then objFso.CreateFolder(strShare & "\Jogos Salvos")
    If Not objFso.FolderExists(strShare & "\Outlook - Assinaturas") Then objFso.CreateFolder(strShare & "\Outlook - Assinaturas")
    If Not objFso.FolderExists(strShare & "\Outlook - Arquivos PST") Then objFso.CreateFolder(strShare & "\Outlook - Arquivos PST")
    If Not objFso.FolderExists(strShare & "\Outlook - App Settings") Then objFso.CreateFolder(strShare & "\Outlook - App Settings")

'Define pasta para o BKP
    strContatS = strShare & "\Contatos"
	strDesktoS = strShare & "\Desktop"
    strMyDocsS = strShare & "\Documentos"
	strDownloS = strShare & "\Downloads"
    strFavoriS = strShare & "\Favoritos"
	strPicturS = strShare & "\Imagens"
	strMusicS = strShare & "\Musicas"
	strVideosS = strShare & "\Videos"
	strGamesS = strShare & "\Jogos Salvos"
    strSignatS = strShare & "\Outlook - Assinaturas"
    strPSTfilS = strShare & "\Outlook - Arquivos PST"
    strAppSetS = strShare & "\Outlook - App Settings"   

'Para definir a origem dos arquivos precisamos primeiro identificar SO
strComputer = "."
Set objWMIService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS in colOSes
OSName =  objOS.Caption
IF OSName = "Microsoft Windows XP Professional" or OSName = "Microsoft Windows 2000 Professional" then
	strContacts = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Contacts"
    strDesktop = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop"
    strMyDocuments = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Meus Documentos"
	strMyDownloads = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Downloads"
    strFavorites = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Favoritos"
	strPictures = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Pictures"
	strMusic = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Music"
	strVideos = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Videos"
	strGames = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Saved Games"
    strSignature = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\assinaturas"
    strPSTfile = objShell.ExpandEnvironmentStrings("%SystemDrive%") & "\outlook\" & strUsername
    strOutSet = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Outlook"
Else
	strContacts = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Contacts"
    strDesktop = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop"
    strMyDocuments = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Documents"
	strMyDownloads = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Downloads"
    strFavorites = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Favorites"
	strPictures = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Pictures"
	strMusic = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Music"
	strVideos = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Videos"
	strGames = objShell.ExpandEnvironmentStrings("%UserProfile%") & "\Saved Games"
    strSignature = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\signatures"
    strPSTfile = objShell.ExpandEnvironmentStrings("%SystemDrive%") & "\outlook\" & strUsername
	strOutSet = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Outlook"
End If

Next

'copia de arquivos

objContacts = RoboCopy & " " & chr(34) & strContacts & chr(34) & " " & Chr(34) & strContatS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objContacts, 1, True

objDesktop = RoboCopy & " " & chr(34) & strDesktop & chr(34) & " " & Chr(34) & strDesktoS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objDesktop, 1, True

objMyDocuments = RoboCopy & " " & chr(34) & strMyDocuments & chr(34) & " " & Chr(34) & strMyDocsS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objMyDocuments, 1, True

objMyDownloads = RoboCopy & " " & chr(34) & strMyDownloads & chr(34) & " " & Chr(34) & strDownloS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objMyDownloads, 1, True

objFavorites = RoboCopy & " " & chr(34) & strFavorites & chr(34) & " " & Chr(34) & strFavoriS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objFavorites, 1, True

objPictures = RoboCopy & " " & chr(34) & strPictures & chr(34) & " " & Chr(34) & strPicturS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objPictures, 1, True

objMusic = RoboCopy & " " & chr(34) & strMusic & chr(34) & " " & Chr(34) & strMusicS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objMusic, 1, True

objVideos = RoboCopy & " " & chr(34) & strVideos & chr(34) & " " & Chr(34) & strVideosS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objVideos, 1, True

objGames = RoboCopy & " " & chr(34) & strGames & chr(34) & " " & Chr(34) & strGamesS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objGames, 1, True

objSignature = RoboCopy & " " & chr(34) & strSignature & chr(34) & " " & Chr(34) & strSignatS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objSignature, 1, True

objOutSet = RoboCopy & " " & chr(34) & strOutSet & chr(34) & " " & Chr(34) & strAppSetS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run objOutSet, 1, True

objPSTfile = RoboCopy & " " & chr(34) & strPSTfile & chr(34) & " " & Chr(34) & strPSTfilS & chr(34) & " " & strCopyOptions & chr(34)
objShell.Run "taskkill /F /IM outlook.exe", 1, True
objShell.Run "taskkill /F /IM outlook.exe", 1, True
objShell.Run objPSTfile, 1, True

