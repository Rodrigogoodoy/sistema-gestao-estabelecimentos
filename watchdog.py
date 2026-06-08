"""
Watchdog: monitora o servidor app.py e reinicia automaticamente se cair.
Roda em background sem janela visivel.
"""
import subprocess
import time
import sys
import os
import urllib.request
from pathlib import Path

BASE_DIR = Path(__file__).parent
PYTHON   = BASE_DIR / ".venv" / "Scripts" / "python.exe"
APP      = BASE_DIR / "app.py"
CHECK_URL = "http://localhost:5000"
CHECK_INTERVAL = 20  # segundos entre verificacoes

os.chdir(BASE_DIR)
os.environ["PYTHONIOENCODING"] = "utf-8"

proc = None

def servidor_vivo():
    try:
        urllib.request.urlopen(CHECK_URL, timeout=5)
        return True
    except:
        return False

def iniciar():
    global proc
    print("Iniciando servidor...")
    proc = subprocess.Popen(
        [str(PYTHON), str(APP)],
        cwd=str(BASE_DIR),
        env=os.environ.copy(),
    )
    print(f"Servidor iniciado (PID {proc.pid})")

iniciar()

while True:
    time.sleep(CHECK_INTERVAL)
    if proc is None or proc.poll() is not None:
        print("Servidor caiu! Reiniciando...")
        time.sleep(3)
        iniciar()
    elif not servidor_vivo():
        print("Servidor nao responde! Reiniciando...")
        try:
            proc.terminate()
            time.sleep(3)
        except:
            pass
        iniciar()
