"""Teste rápido com 1 município para validar as correções."""
import pandas as pd
from playwright.sync_api import sync_playwright
from bot_municipio import processar_municipio

URL = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"

municipio = {"cod_ibge": "4308607", "municipio": "Garibaldi", "uf": "RS"}

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    print(f"Testando {municipio['municipio']}/{municipio['uf']}...")
    resultados = processar_municipio(
        page=page,
        uf=municipio["uf"],
        municipio=municipio["municipio"],
        cod_ibge=municipio["cod_ibge"],
        url_base=URL,
        delay_entre_fichas=0.3,
    )
    browser.close()

print(f"\nTotal coletado: {len(resultados)}")
if resultados:
    df = pd.DataFrame(resultados)
    print("Colunas:", list(df.columns))
    print()
    for campo in ["CNPJ", "Natureza Jurídica(Grupo)", "Data Desativação", "Motivo Desativação"]:
        if campo in df.columns:
            vals = df[campo].fillna("").tolist()
            vazios = sum(1 for v in vals if not str(v).strip())
            print(f"{campo}: {len(vals) - vazios} preenchidos, {vazios} vazios")
            print(f"  Exemplos: {[v for v in vals if str(v).strip()][:3]}")
        else:
            print(f"{campo}: COLUNA AUSENTE")
