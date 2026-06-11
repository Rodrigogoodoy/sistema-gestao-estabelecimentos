import requests, time

s = requests.Session()
s.post('http://localhost:5000/api/login', json={'usuario': 'rodrigo', 'senha': '1234'})

muns = s.get('http://localhost:5000/api/municipios/RS').json()
mapa = {m['nome'].lower(): m for m in muns}

buscar = [
    'Porto Alegre', 'Viamão', 'Canoas', 'São Leopoldo',
    'Novo Hamburgo', 'Cachoeirinha', 'Gravataí', 'Montenegro', 'Santa Cruz do Sul'
]

jobs = []
for nome in buscar:
    m = mapa.get(nome.lower())
    if not m:
        for k, v in mapa.items():
            if nome.lower()[:6] in k:
                m = v
                break
    if m:
        r = s.post('http://localhost:5000/api/pesquisar', json={
            'uf': 'RS', 'municipio': m['nome'], 'cod_ibge': m['cod_ibge'], 'lote': True
        })
        jid = r.json().get('job_id')
        jobs.append((m['nome'], jid))
        print(f"  Disparado: {m['nome']} [{jid}]")
    else:
        print(f"  NAO ENCONTRADO: {nome}")

print(f"\nTotal: {len(jobs)} municípios em andamento...")
print("Aguarde — o Excel combinado abrirá automaticamente quando terminar.")

# Monitora até todos concluirem
while True:
    time.sleep(15)
    concluidos = 0
    erros = 0
    for nome, jid in jobs:
        r = s.get(f'http://localhost:5000/api/status/{jid}').json()
        if r.get('status') == 'done':
            concluidos += 1
        elif r.get('status') == 'error':
            erros += 1
            print(f"  ERRO: {nome}")
    print(f"  Progresso: {concluidos}/{len(jobs)} concluídos...")
    if concluidos + erros == len(jobs):
        print("\nTodos concluídos! Abrindo Excel combinado...")
        job_ids = [jid for _, jid in jobs]
        s.post('http://localhost:5000/api/abrir/lote', json={'job_ids': job_ids})
        break
