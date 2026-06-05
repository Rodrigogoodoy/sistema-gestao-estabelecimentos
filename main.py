import argparse
import json
import pandas as pd
from pathlib import Path
from tqdm import tqdm
from playwright.sync_api import sync_playwright

from config import TARGET_URL, HEADLESS, SAVE_EVERY
from excel_reader import carregar_cnpjs
from bot import processar_cnpj


def _salvar_excel(resultados: list, caminho: Path):
    if not resultados:
        return
    df = pd.DataFrame(resultados)
    colunas_fixas = [c for c in ("CNPJ Pesquisado", "ENCONTRADO") if c in df.columns]
    demais = [c for c in df.columns if c not in colunas_fixas]
    df = df[colunas_fixas + demais]
    df.to_excel(caminho, index=False)


def _ler_checkpoint(path: Path) -> int:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8")).get("ultimo_indice", 0)
        except Exception:
            pass
    return 0


def _salvar_checkpoint(path: Path, indice: int):
    path.write_text(json.dumps({"ultimo_indice": indice}), encoding="utf-8")


def _aguardar_pagina(page):
    page.wait_for_load_state("networkidle", timeout=20_000)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", type=int, default=0,
                        help="Índice do worker (0-based, default=0)")
    parser.add_argument("--workers-total", type=int, default=1,
                        help="Total de workers rodando em paralelo (default=1)")
    parser.add_argument("--limit", type=int, default=None,
                        help="Limita o total de CNPJs (útil para testes)")
    args = parser.parse_args()

    worker_id = args.worker
    workers_total = args.workers_total

    todos_cnpjs = carregar_cnpjs()
    if args.limit:
        todos_cnpjs = todos_cnpjs[:args.limit]
    total_geral = len(todos_cnpjs)

    # Fatia os CNPJs para este worker
    tamanho = (total_geral + workers_total - 1) // workers_total
    inicio_slice = worker_id * tamanho
    fim_slice = min(inicio_slice + tamanho, total_geral)
    cnpjs = todos_cnpjs[inicio_slice:fim_slice]

    checkpoint_file = Path(f"checkpoint_worker_{worker_id}.json")
    output_file = Path(f"resultados_worker_{worker_id}.xlsx")

    inicio = _ler_checkpoint(checkpoint_file)
    pendentes = cnpjs[inicio:]

    sim = nao = erros = 0
    resultados: list[dict] = []

    if inicio > 0 and output_file.exists():
        try:
            df_anterior = pd.read_excel(output_file, dtype=str)
            resultados = df_anterior.to_dict("records")
            sim = len(resultados)  # arquivo só contém SIM
            print(f"[worker {worker_id}] Checkpoint: {sim} resultados anteriores carregados")
        except Exception as e:
            print(f"[worker {worker_id}] Aviso: não foi possível carregar resultados anteriores: {e}")

    print(f"\n[worker {worker_id}] Responsável pelos CNPJs {inicio_slice + 1}–{fim_slice} de {total_geral}")
    print(f"[worker {worker_id}] {len(pendentes)} CNPJs para processar\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.goto(TARGET_URL)
        _aguardar_pagina(page)

        precisa_goto = False  # já estamos na página de busca

        with tqdm(total=len(pendentes), desc=f"Worker {worker_id}", unit="cnpj") as barra:
            for i, cnpj in enumerate(pendentes):
                resultado, precisa_goto = processar_cnpj(page, cnpj, precisa_goto)

                encontrado = resultado.get("ENCONTRADO", "ERRO")
                if encontrado == "SIM":
                    sim += 1
                    resultados.append(resultado)
                elif encontrado == "NÃO":
                    nao += 1
                else:
                    erros += 1

                _salvar_checkpoint(checkpoint_file, inicio + i + 1)
                barra.update(1)
                barra.set_postfix(SIM=sim, NAO=nao, ERRO=erros)

                if (i + 1) % SAVE_EVERY == 0:
                    _salvar_excel(resultados, output_file)

        browser.close()

    _salvar_excel(resultados, output_file)
    print(f"\n[worker {worker_id}] Finalizado!")
    print(f"  Encontrados (SIM) : {sim}")
    print(f"  Não encontrados   : {nao}")
    print(f"  Erros             : {erros}")
    print(f"  Arquivo           : {output_file}")


if __name__ == "__main__":
    main()
