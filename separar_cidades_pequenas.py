"""
Separa os 20.721 estabelecimentos encontrados no CNES por cidades pequenas do Brasil.
Criterio: municipios com populacao ate 50.000 habitantes (IBGE Estimativas 2021).
Saida: pasta "cidades_pequenas/" com um arquivo Excel por UF, cada aba = uma cidade.
"""

import pandas as pd
import os
import re
import json
import gzip
import urllib.request
from collections import defaultdict

ARQUIVO_ENTRADA = "resultados_SIM_final.xlsx"
PASTA_SAIDA = "cidades_pequenas"
LIMITE_POPULACAO = 50_000

# --------------------------------------------------------------------------- #
# 1. Buscar populacao de todos os municipios via API IBGE
# --------------------------------------------------------------------------- #

def buscar_populacao_ibge():
    """Retorna dict {codigo_ibge_6dig: populacao}"""
    print("Buscando dados populacionais do IBGE...")

    for ano in ["2021", "2022"]:
        url = (
            "https://servicodados.ibge.gov.br/api/v3/agregados/6579"
            f"/periodos/{ano}/variaveis/9324"
            "?localidades=N6[all]"
        )
        try:
            req = urllib.request.Request(url, headers={"Accept-Encoding": "identity"})
            with urllib.request.urlopen(req, timeout=60) as resp:
                raw = resp.read()
                # Descompactar gzip se necessario
                if raw[:2] == b'\x1f\x8b':
                    raw = gzip.decompress(raw)
                dados = json.loads(raw.decode("utf-8"))
            break
        except Exception as e:
            print(f"  AVISO: erro ao buscar ano {ano} ({e}), tentando proximo...")
    else:
        raise RuntimeError("Nao foi possivel buscar dados do IBGE. Verifique a conexao.")

    populacao = {}
    for item in dados[0]["resultados"][0]["series"]:
        cod = item["localidade"]["id"]          # 7 digitos
        cod6 = cod[:6]                          # IBGE usa 7 mas nosso dado tem 6
        val = item["serie"].get(ano)
        if val and val != "-":
            try:
                populacao[cod6] = int(val)
            except ValueError:
                pass

    print(f"  {len(populacao)} municipios com dados populacionais carregados.")
    return populacao


# --------------------------------------------------------------------------- #
# 2. Carregar resultados CNES
# --------------------------------------------------------------------------- #

def carregar_dados():
    print(f"Carregando {ARQUIVO_ENTRADA}...")
    df = pd.read_excel(ARQUIVO_ENTRADA)
    print(f"  {len(df)} registros carregados.")

    # Extrair codigo IBGE (6 digitos) e nome da cidade do campo Municipio
    # Formato esperado: "292740 - SALVADOR"
    col_mun = "Município" if "Município" in df.columns else "Municipio"
    df["_ibge6"] = df[col_mun].astype(str).str.extract(r"^(\d{6})")
    df["_cidade"] = df[col_mun].astype(str).str.extract(r"^\d+ - (.+)$")

    sem_cidade = df["_ibge6"].isna().sum()
    if sem_cidade:
        print(f"  AVISO: {sem_cidade} registros sem codigo IBGE valido (serao ignorados).")

    return df


# --------------------------------------------------------------------------- #
# 3. Filtrar cidades pequenas e gerar relatorio
# --------------------------------------------------------------------------- #

def filtrar_cidades_pequenas(df, populacao):
    print(f"\nFiltrando cidades com ate {LIMITE_POPULACAO:,} habitantes...")

    codigos_unicos = df["_ibge6"].dropna().unique()
    pequenas = set()
    sem_pop = []

    for cod in codigos_unicos:
        pop = populacao.get(cod)
        if pop is None:
            sem_pop.append(cod)
        elif pop <= LIMITE_POPULACAO:
            pequenas.add(cod)

    print(f"  Cidades no dataset: {len(codigos_unicos)}")
    print(f"  Cidades pequenas (ate {LIMITE_POPULACAO:,} hab): {len(pequenas)}")
    print(f"  Cidades sem dados de populacao: {len(sem_pop)}")

    df_peq = df[df["_ibge6"].isin(pequenas)].copy()
    print(f"  Registros em cidades pequenas: {len(df_peq)}")

    return df_peq


# --------------------------------------------------------------------------- #
# 4. Salvar arquivos por UF / cidade
# --------------------------------------------------------------------------- #

def sanitizar_nome(nome):
    """Remove caracteres invalidos para nome de aba/arquivo."""
    nome = str(nome).strip()
    nome = re.sub(r'[\\/*?:\[\]]', '', nome)   # caracteres proibidos no Excel
    return nome[:31]                             # limite de 31 chars por aba


def salvar_por_uf(df_peq):
    os.makedirs(PASTA_SAIDA, exist_ok=True)

    # Colunas a exportar (sem as internas _ibge6 e _cidade auxiliares)
    colunas_saida = [c for c in df_peq.columns if not c.startswith("_")]

    ufs = sorted(df_peq["UF"].dropna().unique())
    print(f"\nGerando arquivos para {len(ufs)} estados...")

    resumo = []
    total_registros = 0

    for uf in ufs:
        df_uf = df_peq[df_peq["UF"] == uf]
        cidades = sorted(df_uf["_cidade"].dropna().unique())
        arquivo = os.path.join(PASTA_SAIDA, f"{uf}_cidades_pequenas.xlsx")

        with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
            # Aba de resumo do estado
            resumo_uf = []
            for cidade in cidades:
                df_cidade = df_uf[df_uf["_cidade"] == cidade][colunas_saida]
                aba = sanitizar_nome(cidade)
                df_cidade.to_excel(writer, sheet_name=aba, index=False)
                resumo_uf.append({"Cidade": cidade, "Qtd CNPJs": len(df_cidade)})
                resumo.append({"UF": uf, "Cidade": cidade, "Qtd CNPJs": len(df_cidade)})
                total_registros += len(df_cidade)

            pd.DataFrame(resumo_uf).to_excel(writer, sheet_name="RESUMO", index=False)

        print(f"  {uf}: {len(cidades)} cidades, {len(df_uf)} registros -> {arquivo}")

    # Salvar resumo geral
    arquivo_resumo = os.path.join(PASTA_SAIDA, "_RESUMO_GERAL.xlsx")
    pd.DataFrame(resumo).to_excel(arquivo_resumo, index=False)
    print(f"\nResumo geral salvo em: {arquivo_resumo}")
    print(f"Total de registros exportados: {total_registros}")
    print(f"Arquivos salvos em: {os.path.abspath(PASTA_SAIDA)}/")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    populacao = buscar_populacao_ibge()
    df = carregar_dados()
    df_peq = filtrar_cidades_pequenas(df, populacao)

    if len(df_peq) == 0:
        print("\nNenhum registro encontrado para cidades pequenas.")
    else:
        salvar_por_uf(df_peq)
        print("\nConcluido!")
