@echo off
echo Iniciando 4 workers para verificar Atende SUS...
echo.

start "SUS Worker 0" cmd /k ".\.venv\Scripts\python.exe recheck_sus.py --worker 0 --workers-total 4"
timeout /t 3 /nobreak >nul

start "SUS Worker 1" cmd /k ".\.venv\Scripts\python.exe recheck_sus.py --worker 1 --workers-total 4"
timeout /t 3 /nobreak >nul

start "SUS Worker 2" cmd /k ".\.venv\Scripts\python.exe recheck_sus.py --worker 2 --workers-total 4"
timeout /t 3 /nobreak >nul

start "SUS Worker 3" cmd /k ".\.venv\Scripts\python.exe recheck_sus.py --worker 3 --workers-total 4"

echo.
echo 4 workers iniciados!
echo Quando todos terminarem, rode:
echo   .\.venv\Scripts\python.exe recheck_sus.py --merge
echo.
pause
