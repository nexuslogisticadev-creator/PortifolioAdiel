@echo off
setlocal
set "SERVICE_NAME=FilaComandosExecutorService"
sc stop %SERVICE_NAME%
sc query %SERVICE_NAME%
pause
