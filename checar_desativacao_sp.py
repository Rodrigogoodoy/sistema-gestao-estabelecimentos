import pandas as pd
df = pd.read_excel('SP_MUNICIPIOS_SUS_SIM.xlsx', skiprows=4, dtype=str)
df = df.fillna('')

print(f'Total registros: {len(df)}')
print(f'Colunas: {list(df.columns)}')
print()

if 'Data Desativacao' in df.columns:
    com_data = (df['Data Desativacao'].str.strip() != '').sum()
    sem_data = (df['Data Desativacao'].str.strip() == '').sum()
    print(f'Com Data Desativacao: {com_data}')
    print(f'Sem Data Desativacao (ativos): {sem_data}')
    print()
    print('Exemplos com data:')
    print(df[df['Data Desativacao'].str.strip() != ''][['Razao Social','Municipio','Data Desativacao','Motivo Desativacao']].head(10).to_string())
else:
    print('COLUNA Data Desativacao NAO ENCONTRADA')

# Checar no raw
print()
print('--- Dados brutos (resultados_sp_municipios.xlsx) ---')
df2 = pd.read_excel('resultados_sp_municipios.xlsx', dtype=str).fillna('')
col_desativ = [c for c in df2.columns if 'desativ' in c.lower() or 'Desativ' in c]
print(f'Colunas com desativ: {col_desativ}')
for c in col_desativ:
    preench = (df2[c].str.strip() != '').sum()
    print(f'  {c}: {preench} preenchidos')
