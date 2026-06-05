@echo off
title Sistema CNES DataSUS — Instalacao e Inicio
chcp 65001 >nul
echo.
echo ========================================
echo   Sistema de Gestao de Saude - CNES
echo ========================================
echo.

REM Verifica se Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado no computador.
    echo.
    echo Baixe e instale em: https://www.python.org/downloads/
    echo IMPORTANTE: Marque a opcao "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Python encontrado:
python --version
echo.

REM Cria o ambiente virtual se nao existir
if not exist ".venv" (
    echo Criando ambiente virtual...
    python -m venv .venv
    echo Ambiente virtual criado.
    echo.
)

REM Instala dependencias
echo Instalando dependencias (pode demorar alguns minutos na primeira vez)...
.venv\Scripts\pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERRO ao instalar dependencias. Verifique sua conexao com a internet.
    pause
    exit /b 1
)
echo Dependencias instaladas.
echo.

REM Instala o navegador Chromium do Playwright
echo Instalando navegador Chromium (necessario para o robo)...
.venv\Scripts\python -m playwright install chromium
if errorlevel 1 (
    echo ERRO ao instalar Chromium. Verifique sua conexao com a internet.
    pause
    exit /b 1
)
echo Chromium instalado.
echo.

REM Inicia o servidor
echo ========================================
echo   Iniciando servidor...
echo   Acesse: http://localhost:5000
echo ========================================
echo.
set PYTHONIOENCODING=utf-8
.venv\Scripts\python app.py
pause
