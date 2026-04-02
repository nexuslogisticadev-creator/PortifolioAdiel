@echo off
setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "SERVICE_NAME=FilaComandosExecutorService"
set "EXECUTOR_SCRIPT=%SCRIPT_DIR%executor_fila_comandos.py"
set "NSSM_EXE=%SCRIPT_DIR%nssm\win64\nssm.exe"

set "PYTHON_EXE=%SCRIPT_DIR%.venv314\Scripts\python.exe"

echo.
echo =========================================================
echo  Instalacao do Executor da Fila COMANDOS como Servico
echo =========================================================
echo.

if not exist "%EXECUTOR_SCRIPT%" (
  echo ERRO: executor_fila_comandos.py nao encontrado.
  pause
  goto :END
)

if not exist "%NSSM_EXE%" (
  echo ERRO: NSSM nao encontrado em %NSSM_EXE%
  echo Copie nssm.exe para nssm\win64\ e rode novamente.
  pause
  goto :END
)

echo Usando Python: %PYTHON_EXE%
echo Script Executor: %EXECUTOR_SCRIPT%

echo.
echo Limpando servico anterior...
"%NSSM_EXE%" stop %SERVICE_NAME% >nul 2>&1
"%NSSM_EXE%" remove %SERVICE_NAME% confirm >nul 2>&1

echo Instalando servico...
"%NSSM_EXE%" install %SERVICE_NAME% "%PYTHON_EXE%" "%EXECUTOR_SCRIPT%" >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% AppDirectory "%SCRIPT_DIR%" >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% DisplayName "Executor Fila Comandos" >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% Description "Consumidor externo da aba COMANDOS do Apps Script" >nul 2>&1

if not exist "%SCRIPT_DIR%logs" mkdir "%SCRIPT_DIR%logs"
"%NSSM_EXE%" set %SERVICE_NAME% AppStdout "%SCRIPT_DIR%logs\executor_fila_stdout.log" >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% AppStderr "%SCRIPT_DIR%logs\executor_fila_stderr.log" >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% AppRotateFiles 1 >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% AppExit Default Restart >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% AppRestartDelay 5000 >nul 2>&1
"%NSSM_EXE%" set %SERVICE_NAME% AppEnvironmentExtra "EXECUTOR_ALLOW_LOCAL=false" "EXECUTOR_POLL_SECONDS=8" "EXECUTOR_BATCH_SIZE=25" >nul 2>&1

echo Iniciando servico...
sc start %SERVICE_NAME% >nul 2>&1

echo.
echo OK: Servico %SERVICE_NAME% instalado e iniciado.
echo.
echo Comandos uteis:
echo   sc query %SERVICE_NAME%
echo   sc stop %SERVICE_NAME%
echo   sc start %SERVICE_NAME%
echo.
pause
:END
