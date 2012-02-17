@echo off
echo "Limpando Cache
del c:\ti\hta\logon.hta
del c:\ti\hta\Reparo.hta
copy "\\cemusadobrasil.com.br\NETLOGON\hta\Logon.hta" "c:\ti\hta\Logon.hta" /y
copy "\\cemusadobrasil.com.br\NETLOGON\hta\Reparo.hta" "c:\ti\hta\Reparo.hta" /y
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
echo "Liberando IP"
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /release
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
ECHO Deletando Arquivos Temporários, Cookies, Histórico, Senhas e informações em Formulários
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255
echo "Renovando IP"
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /renew
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 -d \\%COMPUTERNAME% ipconfig /flushdns
gpupdate /force

mshta c:\ti\hta\reparo.hta
ECHO Done!
CLS