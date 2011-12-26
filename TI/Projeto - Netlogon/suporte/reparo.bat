@echo off
echo "Limpando Cache
del c:\ti\hta\logon.hta
nbtstat -R
nbtstat -RR
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
echo "Liberando IP"
ipconfig /release
nbtstat -R
nbtstat -RR
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
echo "Renovando IP"
ipconfig /renew
nbtstat -R
nbtstat -RR
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
ipconfig /flushdns
gpupdate
del c:\ti\hta\logon.hta
ECHO Deletando Arquivos Temporários, Cookies, Histórico, Senhas e informações em Formulários
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32
RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255
ECHO Done!
CLS