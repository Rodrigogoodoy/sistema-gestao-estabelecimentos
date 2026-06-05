"""Corrige CNPJ no resultados_fase2.xlsx: remove CNPJ da mantenedora do campo CNPJ."""
import pandas as pd

df = pd.read_excel("resultados_fase2.xlsx", dtype=str)

def limpar(v):
    v = str(v or "").strip()
    return "" if v in ("---", "-", "nan", "None", "NaN") else v

antes = df["CNPJ"].apply(limpar).ne("").sum()

corrigidos = 0
for i, row in df.iterrows():
    cnpj = limpar(row.get("CNPJ", ""))
    proprio = limpar(row.get("CNPJ Próprio", ""))
    mantenedora = limpar(row.get("CNPJ Mantenedora", ""))

    # Se CNPJ é igual ao da mantenedora e não tem CNPJ Próprio: limpar
    if cnpj and cnpj == mantenedora and not proprio:
        df.at[i, "CNPJ"] = ""
        corrigidos += 1
    elif not cnpj and proprio:
        df.at[i, "CNPJ"] = proprio

depois = df["CNPJ"].apply(limpar).ne("").sum()
print(f"CNPJs antes: {antes} | depois: {depois} | corrigidos: {corrigidos}")

df.to_excel("resultados_fase2.xlsx", index=False)
print("resultados_fase2.xlsx atualizado.")
