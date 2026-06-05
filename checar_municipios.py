import json

with open('municipios.json', encoding='utf-8') as f:
    data = json.load(f)

total = sum(len(v) for v in data.values())
print(f"Total de municipios: {total}")
print(f"Total de estados: {len(data)}")
print()
print("Municipios por estado:")
for uf in sorted(data.keys()):
    print(f"  {uf}: {len(data[uf])}")

# Exemplos de municipios pequenos
print()
print("Exemplo de municipios pequenos em SP:")
sp = sorted(data['SP'], key=lambda x: x['nome'])
for m in sp[:5]:
    print(f"  {m['nome']} (IBGE: {m['cod_ibge']})")
