'Script Para Gera��o de Inventario Basico
'autoria Leonardo Vivas
'Vers�o 2.0
'cria��o 03/06/2009
'modifica��o 03/03/2012
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

'msgbox "N�o parar em caso de erros"
'On Error Resume Next

'msgbox "Carregando variaveis"
varfile = DIR & "\SYS\LOGON.INI"
  Set SYS = oFSO.OpenTextFile(varfile)
  SYSFILE =   SYS.ReadAll
  SYS.close
  execute SYSFILE

'msgbox "Carregando arquivo de Fun��es"
varfile = DIR & "\SYS\FNC.INI"
  Set FNC = oFSO.OpenTextFile(varfile)
  FNCFILE =   FNC.ReadAll
  FNC.close
  execute FNCFILE
  
ChecaArquivoSai(USERLOGS & "\outlook-" & COMP & ".log")

'Bkp Outlook na Profile
objShell.Run "taskkill /F /IM outlook.exe", 0, True
bkpoutlook = "robocopy " & outlookuser &" "& outlookbkpuser & " *.pst " & robosync & LOGUSER &"\outlook-bkp.log"
oShell.Run Robo & " " & DIR & "\HTA\ " & HTA & "\ " & RoboOPSYNC & USERLOGS & "\copyhta.log", 0, False
objShell.Run bkpoutprof, 0, True
objShell.Run bkpoutprof2, 0, True


'Backup outlook
objShell.Run "taskkill /F /IM outlook.exe", 0, True
objShell.Run bkpoutlook, 0, True
