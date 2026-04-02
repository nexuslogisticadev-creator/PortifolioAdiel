@echo off
cd /d "c:\Users\Usuario\Desktop\githubnew"

:start
rem Removemos o /B e o pythonw para a janela ficar VISIVEL
".venv314\Scripts\python.exe" telegram_bot.py

rem Aguarda 10 segundos antes de reiniciar o loop.
rem O loop só continuará se o processo pythonw.exe terminar.
timeout /t 10 /nobreak > nul
goto start
