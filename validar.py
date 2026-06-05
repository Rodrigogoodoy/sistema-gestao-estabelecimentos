import pandas as pd
df = pd.read_excel("MUNICIPIOS_SUS_SIM.xlsx", skiprows=4, dtype=str)
print(f"Total registros: {len(df)}")
print()
campos = ["CNPJ", "Municipio", "Natureza Juridica", "Data Desativacao", "Motivo Desativacao"]
for c in campos:
    if c in df.columns:
        preenchidos = df[c].fillna("").apply(lambda x: x.strip() not in ("", "---")).sum()
        vazios = len(df) - preenchidos
        print(f"{c}: {preenchidos} preenchidos | {vazios} vazios")
    else:
        print(f"{c}: AUSENTE")
print()
print("Municipios e qtd:")
print(df.groupby("Municipio").size().sort_values(ascending=False).to_string())
