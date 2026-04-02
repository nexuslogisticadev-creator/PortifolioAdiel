@echo off
setlocal enabledelayedexpansion
SET SCRIPT_DIR=%~dp0
SET BOT_NAME=TelegramBotService22
SET PYTHON_EXE=%~dp0.venv\Scripts\pythonw.exe
SET BOT_SCRIPT=%SCRIPT_DIR%telegram_bot.py
SET NSSM_DIR=%SCRIPT_DIR%nssm-2.24
SET NSSM_EXE=%NSSM_DIR%\win64\nssm.exe

echo.
echo =========================================================
echo  Configuração do Telegram Bot como Serviço Windows (TURBO)
echo =========================================================
echo.

:: Verifica se o Python.exe existe no ambiente virtual
IF NOT EXIST "%PYTHON_EXE%" (
    echo ❌ ERRO: Python.exe não encontrado em "%PYTHON_EXE%".
    echo Certifique-se de que o ambiente virtual está ativado e o Python existe.
    pause
    GOTO :END
)
echo ✅ Python.exe encontrado: "%PYTHON_EXE%"

:: Verifica se o NSSM já está na pasta (já que baixamos via Chocolatey)
IF NOT EXIST "%NSSM_EXE%" (
    echo ❌ ERRO: NSSM.exe não encontrado em "%NSSM_EXE%".
    echo Por favor, coloque o executável do NSSM na pasta nssm\win64.
    pause
    GOTO :END
)
echo ✅ NSSM.exe encontrado.

:: Parar e remover serviço existente
echo.
echo 🛑 Verificando e parando serviço existente (%BOT_NAME%)...
"%NSSM_EXE%" stop %BOT_NAME% >nul 2>&1
"%NSSM_EXE%" remove %BOT_NAME% confirm >nul 2>&1
echo ✅ Ambiente limpo.

:: Instalando o Serviço
echo.
echo ⚙️ Instalando o bot como um novo serviço Windows (%BOT_NAME%)...
"%NSSM_EXE%" install %BOT_NAME% "%PYTHON_EXE%" "%BOT_SCRIPT%" >nul 2>&1

:: Configurações de Diretório e Nome
echo 📝 Configurando diretório de trabalho...
SET "FIXED_DIR=%SCRIPT_DIR%"
IF "%FIXED_DIR:~-1%"=="\" SET "FIXED_DIR=%FIXED_DIR:~0,-1%"
"%NSSM_EXE%" set %BOT_NAME% AppDirectory "%FIXED_DIR%" >nul 2>&1
"%NSSM_EXE%" set %BOT_NAME% Description "Serviço imortal para o Zé Bot Turbo." >nul 2>&1
"%NSSM_EXE%" set %BOT_NAME% DisplayName "Telegram Bot Serviço" >nul 2>&1

:: Configuração de Logs
echo 📁 Configurando logs...
"%NSSM_EXE%" set %BOT_NAME% AppStdout "%SCRIPT_DIR%logs\%BOT_NAME%_stdout.log" >nul 2>&1
"%NSSM_EXE%" set %BOT_NAME% AppStderr "%SCRIPT_DIR%logs\%BOT_NAME%_stderr.log" >nul 2>&1
"%NSSM_EXE%" set %BOT_NAME% AppRotateFiles 1 >nul 2>&1

:: Configuração de Senha (CRÍTICO)
echo 🔐 Configurando permissões de usuário (7 espaços)...
"%NSSM_EXE%" set %BOT_NAME% ObjectName ".\Usuario" "       " >nul 2>&1

:: Configuração do Modo Imortal (Auto-Recuperação)
echo 🛡️ Configurando auto-recuperação (Modo Imortal)...
"%NSSM_EXE%" set %BOT_NAME% AppExit Default Restart >nul 2>&1
"%NSSM_EXE%" set %BOT_NAME% AppRestartDelay 5000 >nul 2>&1
"%NSSM_EXE%" set %BOT_NAME% AppThrottle 1500 >nul 2>&1

echo ✅ Todas as configurações aplicadas.

:: Iniciando o serviço
echo.
echo 🚀 Iniciando o serviço (%BOT_NAME%)...
sc start %BOT_NAME% >nul 2>&1

echo.
echo =========================================================
echo  ✅ ZÉ BOT TURBO INSTALADO E RODANDO!
echo =========================================================
echo Pressione qualquer tecla para sair.
pause >nul
:END