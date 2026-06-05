import pandas as pd

df = pd.read_excel('resultados_municipios_ativos.xlsx')

col_cnpj = [c for c in df.columns if c == 'CNPJ'][0]
col_proprio = [c for c in df.columns if 'Pr' in c and 'CNPJ' in c][0]
col_mant = [c for c in df.columns if 'Mant' in c and 'CNPJ' in c][0]

INVALIDOS = {'---', '', 'nan', 'NaN', 'None'}

def cnpj_busca(row):
    for col in [col_cnpj, col_proprio, col_mant]:
        val = str(row[col]).strip()
        if val not in INVALIDOS:
            return val
    return ''

df['CNPJ_BUSCA'] = df.apply(cnpj_busca, axis=1)

total = len(df)
com_busca = len(df[df['CNPJ_BUSCA'] != ''])
sem_busca = df[df['CNPJ_BUSCA'] == '']

print(f'Total: {total}')
print(f'Com CNPJ_BUSCA preenchido: {com_busca}')
print(f'Sem nenhum CNPJ (nem mantenedora): {len(sem_busca)}')

if len(sem_busca) > 0:
    print(sem_busca[['Nome', col_cnpj, col_proprio, col_mant]].to_string())

# Reordenar: CNPJ_BUSCA logo apos CNPJ
cols = list(df.columns)
idx = cols.index(col_cnpj)
cols.insert(idx + 1, cols.pop(cols.index('CNPJ_BUSCA')))
df = df[cols]

df.to_excel('resultados_municipios_ativos.xlsx', index=False)
print()
print('Arquivo atualizado: resultados_municipios_ativos.xlsx')
