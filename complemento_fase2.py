"""
Coleta Bento Gonçalves e São Marcos que falharam por acento,
e acrescenta ao resultados_fase2.xlsx já existente.
"""
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright
from bot_municipio import processar_municipio

URL = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"
OUTPUT_FINAL = Path("resultados_fase2.xlsx")

# Nomes sem acento para busca no select (o site armazena sem acento)
PENDENTES = [
    {"posicao": 192, "cod_ibge": "4302105", "municipio": "Bento Goncalves", "uf": "RS"},
    {"posicao": 900, "cod_ibge": "4319000", "municipio": "Sao Marcos",      "uf": "RS"},
]

todos = []
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    for m in PENDENTES:
        print(f"\nMunicípio: {m['municipio']}/{m['uf']}  IBGE: {m['cod_ibge']}")
        try:
            res = processar_municipio(
                page, m["uf"], m["municipio"], m["cod_ibge"], URL
            )
            # Corrige o nome de volta para com acento no resultado
            nome_correto = {"Bento Goncalves": "Bento Gonçalves",
                            "Sao Marcos": "São Marcos"}.get(m["municipio"], m["municipio"])
            for r in res:
                r["Município"] = nome_correto
            todos.extend(res)
            print(f"  Coletados: {len(res)}")
        except Exception as e:
            print(f"  ERRO: {e}")
    browser.close()

if todos:
    df_novo = pd.DataFrame(todos)
    if OUTPUT_FINAL.exists():
        df_ant = pd.read_excel(OUTPUT_FINAL, dtype=str)
        df_final = pd.concat([df_ant, df_novo], ignore_index=True)
    else:
        df_final = df_novo
    df_final.to_excel(OUTPUT_FINAL, index=False)
    print(f"\nSalvo: {OUTPUT_FINAL} ({len(df_final)} registros totais)")
else:
    print("Nenhum resultado coletado.")
