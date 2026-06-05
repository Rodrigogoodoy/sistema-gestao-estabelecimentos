import requests, json
r = requests.get('https://servicodados.ibge.gov.br/api/v1/localidades/municipios?orderBy=nome', timeout=15)
data = r.json()

sem_micro = [m for m in data if not m.get('microrregiao')]
print(f'Sem microrregiao: {len(sem_micro)}')

result = {}
erros = 0
for m in data:
    try:
        uf = m['microrregiao']['mesorregiao']['UF']['sigla']
        if uf not in result:
            result[uf] = []
        result[uf].append({'nome': m['nome'], 'cod_ibge': str(m['id'])})
    except Exception:
        erros += 1

print(f'Erros ao processar: {erros}')
print(f'SP: {len(result.get("SP", []))} municípios')
print(f'RS: {len(result.get("RS", []))} municípios')
print(f'Total estados: {len(result)}')

# Salva como JSON para o site usar sem depender da API
with open('municipios.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)
print('municipios.json salvo com sucesso!')
