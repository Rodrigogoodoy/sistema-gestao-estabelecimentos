"""Teste com os primeiros 5 CNPJs da base real do chefe."""
import pandas as pd
from pathlib import Path
from tqdm import tqdm
from playwright.sync_api import sync_playwright

from config import TARGET_URL, HEADLESS, OUTPUT_FILE
from bot import processar_cnpj
from excel_reader import carregar_cnpjs

LIMITE = 1000
OUTPUT_TESTE = Path("resultados_teste_real.xlsx")

def _salvar_excel(resultados, caminho):
    if not resultados:
        return
    df = pd.DataFrame(resultados)
    colunas_fixas = [c for c in ("CNPJ Pesquisado", "ENCONTRADO") if c in df.columns]
    demais = [c for c in df.columns if c not in colunas_fixas]
    df = df[colunas_fixas + demais]
    df.to_excel(caminho, index=False)
    print(f"\nResultados salvos em: {caminho}")

def main():
    todos = carregar_cnpjs()
    cnpjs = todos[:LIMITE]
    print(f"Testando os primeiros {LIMITE} CNPJs da base real\n")

    resultados = []
    sim = nao = erros = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.goto(TARGET_URL)

        precisa_goto = False  # já estamos na página de busca após page.goto acima
        with tqdm(total=len(cnpjs), desc="Progresso", unit="cnpj") as barra:
            for cnpj in cnpjs:
                resultado, precisa_goto = processar_cnpj(page, cnpj, precisa_goto)
                resultados.append(resultado)
                encontrado = resultado.get("ENCONTRADO", "ERRO")
                if encontrado == "SIM":
                    sim += 1
                elif encontrado == "NÃO":
                    nao += 1
                else:
                    erros += 1
                barra.update(1)
                barra.set_postfix(SIM=sim, NAO=nao, ERRO=erros)

        browser.close()

    _salvar_excel(resultados, OUTPUT_TESTE)
    print(f"\nFinalizado!")
    print(f"  Encontrados (SIM): {sim}")
    print(f"  Não encontrados:   {nao}")
    print(f"  Erros:             {erros}")

    print("\nDetalhes por CNPJ:")
    for r in resultados:
        cnpj = r.get("CNPJ Pesquisado", "?")
        status = r.get("ENCONTRADO", "?")
        campos = len([k for k in r if k not in ("CNPJ Pesquisado", "ENCONTRADO")])
        if status == "SIM":
            print(f"  {cnpj} -> {status} ({campos} campos coletados)")
        else:
            print(f"  {cnpj} -> {status}")

if __name__ == "__main__":
    main()
