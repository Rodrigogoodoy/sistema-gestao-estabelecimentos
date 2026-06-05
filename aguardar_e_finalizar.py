"""
Fica em segundo plano monitorando os 4 workers.
Quando todos terminarem, executa merge_results.py automaticamente.
"""
import json
import time
import subprocess
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd

WORKERS_TOTAL = 4
CHECK_INTERVAL = 60  # segundos entre cada verificação
LOG = Path("log_monitor_final.txt")


def log(msg):
    ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    linha = f"[{ts}] {msg}"
    print(linha, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(linha + "\n")


def tamanhos_workers():
    cnpjs = pd.read_excel("050526__Base_Pesquisa_DataSUS.xlsx", dtype=str)["cnpj"].dropna().tolist()
    total = len(cnpjs)
    tamanho = (total + WORKERS_TOTAL - 1) // WORKERS_TOTAL
    tamanhos = []
    for w in range(WORKERS_TOTAL):
        inicio = w * tamanho
        fim = min(inicio + tamanho, total)
        tamanhos.append(fim - inicio)
    return tamanhos


def worker_terminou(worker_id, tamanho_slice):
    cp = Path(f"checkpoint_worker_{worker_id}.json")
    if not cp.exists():
        return False
    try:
        data = json.loads(cp.read_text(encoding="utf-8"))
        return data.get("ultimo_indice", 0) >= tamanho_slice
    except Exception:
        return False


def main():
    log("Monitor iniciado. Aguardando os 4 workers terminarem...")
    log(f"Verificando a cada {CHECK_INTERVAL} segundos.")

    tamanhos = tamanhos_workers()
    log(f"Tamanhos dos slices: {tamanhos}")

    while True:
        status = [worker_terminou(i, tamanhos[i]) for i in range(WORKERS_TOTAL)]
        prontos = sum(status)
        log(f"Progresso: {prontos}/{WORKERS_TOTAL} workers concluídos {status}")

        if all(status):
            log("Todos os workers terminaram! Executando merge_results.py...")
            resultado = subprocess.run(
                [sys.executable, "merge_results.py"],
                capture_output=True, text=True, encoding="utf-8"
            )
            log("=== Saída do merge_results.py ===")
            for linha in resultado.stdout.splitlines():
                log(linha)
            if resultado.returncode == 0:
                log("CONCLUÍDO! Arquivos gerados: resultados_final.xlsx e ESTABELECIMENTOS_CNES.xlsx")
            else:
                log(f"ERRO no merge: {resultado.stderr}")
            break

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
