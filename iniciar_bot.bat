@echo off
REM Script para iniciar o Bot Telegram permanentemente
REM Elimina processos Python existentes e inicia o bot limpo
taskkill /F /IM python.exe 2>nul
timeout /t 3 /nobreak
cd /d "%~dp0"
del bot.lock 2>nul
echo Iniciando Bot Telegram...
".venv314\Scripts\python.exe" telegram_bot.py
pause
