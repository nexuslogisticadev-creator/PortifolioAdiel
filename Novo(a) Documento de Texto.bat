@echo off
echo Finalizando Bot, Robo e o Loop do BAT...
taskkill /F /IM pythonw.exe /T
taskkill /F /IM python.exe /T
taskkill /F /IM cmd.exe /FI "WINDOWTITLE eq c:\windows\system32\cmd.exe" /T
echo.
echo Todos os processos foram encerrados com sucesso.
pause
