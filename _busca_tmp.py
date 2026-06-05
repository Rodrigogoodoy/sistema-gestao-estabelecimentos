import json, unicodedata

def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper()

with open("municipios.json", encoding="utf-8") as f:
    dados = json.load(f)

# Busca CONTENDA no PR
pr = dados.get("PR", [])
sc = dados.get("SC", [])

for m in pr:
    if "CONTENDA" in sem_acento(m["nome"]) or "COTENDA" in sem_acento(m["nome"]):
        print(f"PR match contenda: {m['nome']} | {m['cod_ibge']}")

for m in sc:
    if "NEGRINHO" in sem_acento(m["nome"]) or "BENTO DO SUL" in sem_acento(m["nome"]):
        print(f"SC match: {m['nome']} | {m['cod_ibge']}")

for m in pr:
    if "NEGRINHO" in sem_acento(m["nome"]) or "BENTO DO SUL" in sem_acento(m["nome"]):
        print(f"PR match: {m['nome']} | {m['cod_ibge']}")
