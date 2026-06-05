import pandas as pd
df = pd.read_excel('CAXIAS_SUS_SIM.xlsx', skiprows=4, dtype=str)
df = df.fillna('')

sem_cnpj = df[df['CNPJ'].str.strip() == '']
com_cnpj = df[df['CNPJ'].str.strip() != '']

print(f'Total: {len(df)}')
print(f'Sem CNPJ: {len(sem_cnpj)}')
print(f'Com CNPJ: {len(com_cnpj)}')
print()
print('--- Sem CNPJ ---')
ativos_sem = (sem_cnpj['Data Desativacao'].str.strip() == '').sum()
inativos_sem = (sem_cnpj['Data Desativacao'].str.strip() != '').sum()
print(f'Ativos (sem data desativacao): {ativos_sem}')
print(f'Com data desativacao: {inativos_sem}')
print()
print('--- Com CNPJ ---')
ativos_com = (com_cnpj['Data Desativacao'].str.strip() == '').sum()
inativos_com = (com_cnpj['Data Desativacao'].str.strip() != '').sum()
print(f'Ativos: {ativos_com}')
print(f'Com data desativacao: {inativos_com}')
