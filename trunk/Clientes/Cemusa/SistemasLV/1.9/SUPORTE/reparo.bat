@echo off
rmdir c:\ti\hta /S /Q
mkdir c:\ti\hta
copy "\\10.10.1.2\hta\Reparo.hta" "c:\ti\hta\Reparo.hta" /y
copy "\\10.10.1.2\hta\Logon.hta" "c:\ti\hta\Logon.hta" /y
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 netsh firewall set opmode disable
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /release & ipconfig /renew
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 nbtstat -R
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 nbtstat -RR
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
\\csrv06\ti$\pstools\psexec.exe -h -u cemusa\informatica -p 654321 ipconfig /flushdns
gpupdate
gpupdate /force
mshta c:\ti\hta\reparo.hta
rmdir c:\ti\hta /S /Q