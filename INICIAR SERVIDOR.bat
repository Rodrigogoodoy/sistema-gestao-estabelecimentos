@echo off
title Servidor do Site - Rede Local
echo Iniciando servidor...
powershell -ExecutionPolicy Bypass -File "%~dp0servidor.ps1"
pause
