'Script do logon
'autoria Leonardo Vivas
'Versão 1.0
'criação 03/06/2009
'modificação 08/02/2011
' -----------------------------------------------------------------' 
Set objnet = CreateObject("WScript.Network")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Não parar em caso de erros
'On Error Resume Next

strdiasemana = Weekday(now)
robocopy = "c:\suporte\robocopy.exe C:\Users\administrador\.VirtualBox\ d:\bkp-srv\VirtualBox\TEMP\ /TEE /S /E /COPY:DAT /R:10 /W:45 /XF *.db /LOG+:c:\logs\VM.log"
TS = "c:\progra~1\Oracle\VirtualBox\VBoxManage.exe startvm --type gui 5881262c-e4a0-4ee2-a4d7-0a3f5972c8c5"
FTP = "c:\progra~1\Oracle\VirtualBox\VBoxManage.exe startvm --type gui 9d0e88fa-b9bb-4211-b0e5-924fd6d2cb62"
CSRV04 = "c:\progra~1\Oracle\VirtualBox\VBoxManage.exe startvm --type gui 16213085-e03b-4943-a48a-8f0579b89287"

Select Case strdiasemana
Case 0
strdiasemana = "Sábado"
Case 1
strdiasemana = "Domingo"
Case 2
strdiasemana = "Segunda"
Case 3
strdiasemana = "Terça"
Case 4
strdiasemana = "Quarta"
Case 5
strdiasemana = "Quinta"
Case 6
strdiasemana = "Sexta"
End Select

'Apaga a semana Anterios
objFSO.DeleteFile ("d:\bkp-srv\TS-BACKUP-"& strdiasemana & "*.*")
'para maquinas virtuais
objshell.run FTP, 0, True
objshell.run CSRV04, 0, True
objshell.run TS, 0, True
'Copia Arquivos
objshell.run robocopy , 0, True
'Compacta
objshell.run "c:\suporte\rar.exe a -t -df  -agdd-mmm-yy d:\bkp-srv\TS-BACKUP-"& strdiasemana & "-.rar d:\bkp-srv\VirtualBox\TEMP\*.*", 1, True
'Reinicia as VM's
SHUTDOWN -f -r -t 300 -c "Seu computador sera Desligado Dentro de 5 Minutos para manutenção."