Set ObjShell = CreateObject("wscript.shell")
strShutdown = "SHUTDOWN -r -t 300 -c ""Seu Computador sera Reiniciado Dentro de 5 Minutos"" "
ObjShell.run strShutdown
