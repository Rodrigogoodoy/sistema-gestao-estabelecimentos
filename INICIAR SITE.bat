@echo off
chcp 65001 > nul
title Sistema CNES DataSUS
color 1F
echo.
echo  =========================================
echo    INICIANDO SISTEMA CNES DataSUS...
echo  =========================================
echo.
echo  Aguarde enquanto o servidor inicia...
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
.venv\Scripts\python.exe app.py
echo.
echo  Servidor encerrado.
pause
