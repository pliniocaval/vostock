@echo off
del c:\ti\hta\logon.hta
del c:\ti\hta\Reparo.hta
copy "\\cemusadobrasil.com.br\NETLOGON\hta\Logon.hta" "c:\ti\hta\Logon.hta" /y
copy "\\cemusadobrasil.com.br\NETLOGON\hta\Reparo.hta" "c:\ti\hta\Reparo.hta" /y
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /release
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 netsh firewall set opmode disable
ipconfig /renew
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -u cemusa\informatica -p 654321 ipconfig /flushdns
gpupdate /force
gpupdate /force
gpupdate /force
mshta c:\ti\hta\reparo.hta