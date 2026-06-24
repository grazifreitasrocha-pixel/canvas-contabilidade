@echo off
title Canvas Bot
cd /d "%~dp0"

:loop
echo [%date% %time%] Iniciando Canvas Bot...
.venv\Scripts\python.exe canvas_bot.py
set EXIT=%ERRORLEVEL%

if %EXIT%==1 (
    echo [%date% %time%] Outra instancia ja esta rodando. Nao reiniciando.
    pause
    exit /b 1
)

echo [%date% %time%] Bot parou. Reiniciando em 10 segundos...
timeout /t 10 /nobreak
goto loop
