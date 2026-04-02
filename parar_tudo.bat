@echo off
echo Finalizando Bot, Robo e o Loop do BAT...
REM Cria o arquivo STOP_ALL para encerrar todos os robôs de forma limpa
echo Parada global > STOP_ALL
timeout /t 2 > nul
taskkill /F /IM pythonw.exe /T
taskkill /F /IM python.exe /T
taskkill /F /IM cmd.exe /FI "WINDOWTITLE eq c:\windows\system32\cmd.exe" /T
echo.
echo Todos os processos foram encerrados com sucesso.
pause