@echo off
title ROBO CNES - Processamento Fim de Semana
color 0A
echo.
echo  ============================================================
echo   ROBO CNES - 4 WORKERS PARALELOS
echo   Iniciando em: %date% %time%
echo   136.000 CNPJs - Previsao: ~40 horas
echo  ============================================================
echo.
echo  Abrindo 4 janelas de processamento...
echo  NAO feche nenhuma janela ate terminar!
echo.

cd /d "%~dp0"

set PYTHON=%~dp0.venv\Scripts\python.exe
set PYTHONIOENCODING=utf-8

start "CNES Worker 0 (CNPJs 1-34000)" cmd /k "%PYTHON% main.py --worker 0 --workers-total 4"
timeout /t 5 /nobreak >nul

start "CNES Worker 1 (CNPJs 34001-68000)" cmd /k "%PYTHON% main.py --worker 1 --workers-total 4"
timeout /t 5 /nobreak >nul

start "CNES Worker 2 (CNPJs 68001-102000)" cmd /k "%PYTHON% main.py --worker 2 --workers-total 4"
timeout /t 5 /nobreak >nul

start "CNES Worker 3 (CNPJs 102001-136000)" cmd /k "%PYTHON% main.py --worker 3 --workers-total 4"

echo.
echo  4 workers iniciados com sucesso!
echo.
echo  Para acompanhar o progresso, abra um novo terminal e execute:
echo      %PYTHON% status.py
echo.
echo  Quando todas as janelas mostrarem "Finalizado!", execute:
echo      %PYTHON% merge_results.py
echo.
pause
