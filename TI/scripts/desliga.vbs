Set ObjShell = CreateObject("wscript.shell")
strShutdown = "SHUTDOWN -s -t 100 -c ""Seu computador sera Desligado Dentro de 2 Minutos"" "
ObjShell.run strShutdown
