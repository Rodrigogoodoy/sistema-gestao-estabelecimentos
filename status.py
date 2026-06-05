"""
Mostra o progresso dos workers em tempo real.
Execute a qualquer momento: python status.py
"""
import sys
import json
from pathlib import Path
from datetime import datetime, timedelta

# Garante UTF-8 no terminal Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')


WORKERS_TOTAL = 4
SEGUNDOS_POR_CNPJ = 4.0  # estimativa por worker


def status():
    # Tenta descobrir o total de CNPJs a partir dos checkpoints
    total_geral = 0
    try:
        from excel_reader import carregar_cnpjs
        todos = carregar_cnpjs()
        total_geral = len(todos)
    except Exception:
        total_geral = 136000  # fallback

    tamanho = (total_geral + WORKERS_TOTAL - 1) // WORKERS_TOTAL

    print(f"\n{'='*52}")
    print(f"  STATUS DO ROBÔ CNES — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"  Total de CNPJs: {total_geral:,}")
    print(f"{'='*52}")

    total_feito = 0
    for w in range(WORKERS_TOTAL):
        inicio_slice = w * tamanho
        fim_slice = min(inicio_slice + tamanho, total_geral)
        cnpjs_worker = fim_slice - inicio_slice

        ck_file = Path(f"checkpoint_worker_{w}.json")
        feito = 0
        if ck_file.exists():
            try:
                feito = json.loads(ck_file.read_text(encoding="utf-8")).get("ultimo_indice", 0)
            except Exception:
                pass

        pct = feito / cnpjs_worker * 100 if cnpjs_worker else 0
        blocos = int(pct / 5)
        barra = "█" * blocos + "░" * (20 - blocos)

        status_txt = "✅ Concluído" if feito >= cnpjs_worker else f"{feito:>6,}/{cnpjs_worker:,}"
        print(f"  Worker {w}: [{barra}] {pct:5.1f}%  {status_txt}")
        total_feito += feito

    pct_total = total_feito / total_geral * 100 if total_geral else 0
    restantes = total_geral - total_feito

    print(f"{'─'*52}")
    blocos = int(pct_total / 5)
    barra = "█" * blocos + "░" * (20 - blocos)
    print(f"  TOTAL  : [{barra}] {pct_total:5.1f}%  {total_feito:,}/{total_geral:,}")

    if total_feito > 0 and restantes > 0:
        # Com 4 workers paralelos
        seg_restantes = (restantes / WORKERS_TOTAL) * SEGUNDOS_POR_CNPJ
        eta = datetime.now() + timedelta(seconds=seg_restantes)
        print(f"  ETA    : ~{eta.strftime('%d/%m (%a) às %H:%M')}")
    elif restantes == 0 and total_feito > 0:
        print(f"  STATUS : 🎉 PROCESSAMENTO CONCLUÍDO!")
        print(f"  Execute: python merge_results.py")

    print(f"{'='*52}\n")


if __name__ == "__main__":
    status()
