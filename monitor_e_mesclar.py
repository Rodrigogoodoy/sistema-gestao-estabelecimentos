"""
Monitora os workers e executa merge_results.py automaticamente quando todos terminarem.
Deixe essa janela aberta junto com os workers.
"""
import json
import time
import subprocess
import sys
from pathlib import Path
from datetime import datetime

WORKERS_TOTAL = 4
INTERVALO_VERIFICACAO = 300  # checa a cada 5 minutos


def todos_concluidos(total_geral: int) -> bool:
    tamanho = (total_geral + WORKERS_TOTAL - 1) // WORKERS_TOTAL
    for w in range(WORKERS_TOTAL):
        inicio = w * tamanho
        fim = min(inicio + tamanho, total_geral)
        cnpjs_worker = fim - inicio
        ck = Path(f"checkpoint_worker_{w}.json")
        feito = 0
        if ck.exists():
            try:
                feito = json.loads(ck.read_text(encoding="utf-8")).get("ultimo_indice", 0)
            except Exception:
                pass
        if feito < cnpjs_worker:
            return False
    return True


def progresso(total_geral: int) -> str:
    tamanho = (total_geral + WORKERS_TOTAL - 1) // WORKERS_TOTAL
    total_feito = 0
    for w in range(WORKERS_TOTAL):
        ck = Path(f"checkpoint_worker_{w}.json")
        if ck.exists():
            try:
                total_feito += json.loads(ck.read_text(encoding="utf-8")).get("ultimo_indice", 0)
            except Exception:
                pass
    pct = total_feito / total_geral * 100 if total_geral else 0
    return f"{total_feito:,}/{total_geral:,} ({pct:.1f}%)"


def main():
    try:
        from excel_reader import carregar_cnpjs
        total_geral = len(carregar_cnpjs())
    except Exception:
        total_geral = 136000

    print(f"\n{'='*52}")
    print(f"  MONITOR AUTOMÁTICO — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"  Aguardando {WORKERS_TOTAL} workers finalizarem...")
    print(f"  Verificando a cada {INTERVALO_VERIFICACAO // 60} minutos")
    print(f"{'='*52}\n")

    while True:
        agora = datetime.now().strftime("%d/%m %H:%M")
        prog = progresso(total_geral)
        print(f"[{agora}] Progresso: {prog}")

        if todos_concluidos(total_geral):
            print(f"\n[{agora}] TODOS OS WORKERS CONCLUÍDOS!")
            print("  Executando merge_results.py...\n")
            python = sys.executable
            resultado = subprocess.run([python, "merge_results.py"], cwd=str(Path(__file__).parent))
            if resultado.returncode == 0:
                print("\n" + "="*52)
                print("  CONCLUÍDO COM SUCESSO!")
                print("  Arquivo gerado: ESTABELECIMENTOS_CNES.xlsx")
                print("="*52)
            else:
                print("\nERRO ao executar merge_results.py — execute manualmente.")
            break

        time.sleep(INTERVALO_VERIFICACAO)


if __name__ == "__main__":
    main()
