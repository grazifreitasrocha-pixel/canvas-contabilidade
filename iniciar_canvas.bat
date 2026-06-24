@echo off
title Canvas Bot
cd /d "%~dp0"
set PYTHONUTF8=1

if not exist .venv (
    echo Criando ambiente virtual...
    uv venv .venv
)

echo Instalando dependencias...
uv pip install -r requirements.txt --python .venv\Scripts\python.exe 2>&1

echo.
echo Iniciando Canvas Bot...
.venv\Scripts\python.exe canvas_bot.py > log.txt 2>&1

echo.
type log.txt
pause
