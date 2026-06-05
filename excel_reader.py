import re
import pandas as pd
from config import EXCEL_FILE, EXCEL_SHEET, EXCEL_COLUMN


def _limpar_cnpj(valor: str) -> str:
    return re.sub(r"\D", "", str(valor))


def carregar_cnpjs() -> list[str]:
    df = pd.read_excel(EXCEL_FILE, sheet_name=EXCEL_SHEET, dtype=str)

    if EXCEL_COLUMN not in df.columns:
        colunas = list(df.columns)
        raise ValueError(
            f"Coluna '{EXCEL_COLUMN}' não encontrada. Colunas disponíveis: {colunas}"
        )

    cnpjs = (
        df[EXCEL_COLUMN]
        .dropna()
        .map(_limpar_cnpj)
        .map(lambda x: x.zfill(14))            # garante 14 dígitos com zeros à esquerda
        .pipe(lambda s: s[s.str.len() == 14])  # descarta qualquer valor inválido
        .tolist()
    )

    print(f"[excel_reader] {len(cnpjs)} CNPJs válidos carregados de '{EXCEL_FILE}'")
    return cnpjs
