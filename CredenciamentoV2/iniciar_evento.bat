@echo off
chcp 65001 >nul
title Credenciamento do Evento - Iniciando...

REM ============================================================
REM  Este script sobe o servidor (Flask) e o tunel HTTPS (ngrok)
REM  automaticamente, em duas janelas separadas.
REM
REM  ANTES DE USAR:
REM  1) Este arquivo deve estar dentro da pasta "evento_app"
REM     (na mesma pasta do arquivo app.py)
REM  2) Ajuste a linha NGROK_PATH abaixo com o caminho do seu
REM     ngrok.exe (o padrao sugerido e' C:\ngrok\ngrok.exe)
REM  3) Se ainda nao autenticou o ngrok, rode uma vez antes:
REM     ngrok config add-authtoken SEU_TOKEN_AQUI
REM ============================================================

set NGROK_PATH=C:\ngrok\ngrok.exe

echo.
echo ================================================
echo   Credenciamento do Evento
echo ================================================
echo.

REM Vai para a pasta onde este .bat esta salvo (a pasta do projeto)
cd /d "%~dp0"

echo [1/2] Iniciando o servidor local (Flask)...
start "Servidor - Credenciamento" cmd /k "python app.py"

echo Aguardando o servidor subir...
timeout /t 3 /nobreak >nul

echo [2/2] Iniciando o tunel HTTPS (ngrok)...
if exist "%NGROK_PATH%" (
    start "Tunel HTTPS - ngrok" cmd /k ""%NGROK_PATH%" http 5000"
) else (
    echo.
    echo   ATENCAO: nao encontrei o ngrok em "%NGROK_PATH%"
    echo   Edite este arquivo (clique direito - Editar) e corrija
    echo   a linha NGROK_PATH com o caminho correto do ngrok.exe.
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  Duas janelas foram abertas:
echo   1) "Servidor - Credenciamento"  -^> nao feche
echo   2) "Tunel HTTPS - ngrok"        -^> pegue o link
echo      que aparece na linha "Forwarding" (https://...)
echo      e use esse link no celular do recepcionista,
echo      terminando em /scanner
echo      Ex: https://xxxx.ngrok-free.app/scanner
echo ================================================
echo.
pause
