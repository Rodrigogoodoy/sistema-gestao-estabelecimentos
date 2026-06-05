import pandas as pd

df = pd.read_excel('SP_MUNICIPIOS_SUS_SIM.xlsx', skiprows=4, dtype=str)
df = df.fillna('')

print("=" * 60)
print("CONFERENCIA COMPLETA — SP_MUNICIPIOS_SUS_SIM.xlsx")
print("=" * 60)

# 1. Total e municípios
print(f"\n[1] TOTAL DE REGISTROS: {len(df)}")
print("\nPor município:")
print(df.groupby('Municipio').size().sort_values(ascending=False).to_string())

# 2. Colunas presentes
print("\n[2] COLUNAS PRESENTES:")
for c in df.columns:
    print(f"  - {c}")

# 3. Data Desativação
print("\n[3] DATA DESATIVACAO:")
com_data   = (df['Data Desativacao'].str.strip() != '').sum()
sem_data   = (df['Data Desativacao'].str.strip() == '').sum()
print(f"  Ativos (sem data): {sem_data}")
print(f"  Com data (desativados): {com_data}")
invalidos  = df['Data Desativacao'].str.contains(r'null|---|\{\{', na=False).sum()
print(f"  Valores inválidos (null/---): {invalidos}")

# 4. Motivo Desativação
print("\n[4] MOTIVO DESATIVACAO:")
com_motivo = (df['Motivo Desativacao'].str.strip() != '').sum()
print(f"  Com motivo: {com_motivo}")
print(f"  Sem motivo (ativos): {len(df) - com_motivo}")

# 5. Natureza Jurídica
print("\n[5] NATUREZA JURIDICA:")
nj = df['Natureza Juridica'].str.strip()
preench = (nj != '').sum()
vazio   = (nj == '').sum()
invalido = nj.str.contains(r'null|---|\{\{', na=False).sum()
print(f"  Preenchidos: {preench}")
print(f"  Vazios: {vazio}")
print(f"  Inválidos: {invalido}")
print("  Valores distintos:")
for v, c in nj.value_counts().items():
    if v.strip():
        print(f"    [{c:>4}] {v}")

# 6. CNPJ
print("\n[6] CNPJ:")
cnpj = df['CNPJ'].str.strip()
com_cnpj  = (cnpj != '').sum()
sem_cnpj  = (cnpj == '').sum()
invalido_cnpj = cnpj.str.contains(r'null|---|\{\{', na=False).sum()
print(f"  Com CNPJ próprio: {com_cnpj}")
print(f"  Sem CNPJ (públicos municipais): {sem_cnpj}")
print(f"  Inválidos: {invalido_cnpj}")

# 7. Campos críticos com valores ruins
print("\n[7] VALORES INVALIDOS (null/---/{{) NOS CAMPOS CRITICOS:")
campos_criticos = ['CNPJ', 'Razao Social', 'Tipo de Estabelecimento',
                   'Natureza Juridica', 'CEP', 'Municipio', 'UF']
ok = True
for c in campos_criticos:
    if c in df.columns:
        bad = df[c].str.contains(r'null|---|\{\{', na=False).sum()
        if bad > 0:
            print(f"  ATENCAO — {c}: {bad} valores inválidos")
            ok = False
if ok:
    print("  Nenhum valor inválido encontrado. OK!")

# 8. Registros duplicados (por Codigo CNES)
print("\n[8] DUPLICATAS (por Codigo CNES):")
dupl = df[df.duplicated('Codigo CNES', keep=False) & (df['Codigo CNES'].str.strip() != '')]
print(f"  Registros duplicados: {len(dupl)}")
if len(dupl) > 0:
    print(dupl[['Codigo CNES','Razao Social','Municipio']].to_string())

# 9. Amostra dos desativados
print("\n[9] AMOSTRA — DESATIVADOS (primeiros 10):")
desativ = df[df['Data Desativacao'].str.strip() != ''][
    ['Razao Social','Municipio','Data Desativacao','Motivo Desativacao']
].head(10)
print(desativ.to_string())

# 10. Resumo final
print("\n" + "=" * 60)
print("RESUMO FINAL")
print("=" * 60)
print(f"  Total estabelecimentos: {len(df)}")
print(f"  Ativos: {sem_data}")
print(f"  Desativados: {com_data}")
print(f"  Municípios: {df['Municipio'].nunique()}")
print(f"  UF: {df['UF'].unique()}")
print("=" * 60)
