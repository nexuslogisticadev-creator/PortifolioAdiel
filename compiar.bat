@echo off
REM Script para compilar PAINEL com PyInstaller
REM Com logging configurado para funcionar em .exe

echo.
echo ====================================
echo COMPILANDO PAINEL.EXE COM LOGGING
echo ====================================
echo.

REM Limpar build anterior
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist painel.spec del painel.spec

REM Compilar
pyinstaller --onefile --noconsole ^
  --add-data "config.json:." ^
  --hidden-import=pandas ^
  --hidden-import=openpyxl ^
  --hidden-import=customtkinter ^
  --name painel ^
  painel.py

echo.
echo ====================================
if exist dist\painel.exe (
  echo ✅ PAINEL.EXE criado com sucesso!
  echo 📁 Localização: %cd%\dist\painel.exe
  echo.
  echo 📝 Logs salvos em: painel_debug.log
  echo    (na mesma pasta do executável)
) else (
  echo ❌ ERRO: painel.exe não foi criado
)
echo ====================================
echo.
pause
