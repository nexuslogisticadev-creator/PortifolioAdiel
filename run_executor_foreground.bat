@echo off
setlocal
cd /d "%~dp0"

if exist ".venv314\Scripts\python.exe" (
  .venv314\Scripts\python.exe executor_fila_comandos.py
  goto :END
)

echo Ambiente .venv314 não encontrado!
pause
goto :END

:END
pause
