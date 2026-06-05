import pandas as pd

df = pd.read_excel('CAXIAS_SUS_SIM.xlsx', skiprows=4, dtype=str)
df = df.fillna('')

print("=" * 60)
print("CONFERENCIA COMPLETA — CAXIAS_SUS_SIM.xlsx")
print("=" * 60)

print(f"\n[1] TOTAL DE REGISTROS: {len(df)}")

print("\n[2] COLUNAS PRESENTES:")
for c in df.columns:
    print(f"  - {c}")

print("\n[3] DATA DESATIVACAO:")
com_data = (df['Data Desativacao'].str.strip() != '').sum()
sem_data = (df['Data Desativacao'].str.strip() == '').sum()
invalidos = df['Data Desativacao'].str.contains(r'null|---|\{\{', na=False).sum()
print(f"  Ativos (sem data): {sem_data}")
print(f"  Com data (desativados): {com_data}")
print(f"  Valores inválidos: {invalidos}")

print("\n[4] MOTIVO DESATIVACAO:")
com_motivo = (df['Motivo Desativacao'].str.strip() != '').sum()
print(f"  Com motivo: {com_motivo}")
print(f"  Sem motivo (ativos): {len(df) - com_motivo}")

print("\n[5] NATUREZA JURIDICA:")
nj = df['Natureza Juridica'].str.strip()
preench  = (nj != '').sum()
vazio    = (nj == '').sum()
invalido = nj.str.contains(r'null|---|\{\{', na=False).sum()
print(f"  Preenchidos: {preench}")
print(f"  Vazios: {vazio}")
print(f"  Inválidos: {invalido}")
print("  Valores distintos:")
for v, c in nj.value_counts().items():
    if v.strip():
        print(f"    [{c:>4}] {v}")

print("\n[6] CNPJ:")
cnpj = df['CNPJ'].str.strip()
com_cnpj      = (cnpj != '').sum()
sem_cnpj      = (cnpj == '').sum()
invalido_cnpj = cnpj.str.contains(r'null|---|\{\{', na=False).sum()
print(f"  Com CNPJ próprio: {com_cnpj}")
print(f"  Sem CNPJ (públicos municipais): {sem_cnpj}")
print(f"  Inválidos: {invalido_cnpj}")

print("\n[7] VALORES INVALIDOS NOS CAMPOS CRITICOS:")
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

print("\n[8] DUPLICATAS (por Codigo CNES):")
dupl = df[df.duplicated('Codigo CNES', keep=False) & (df['Codigo CNES'].str.strip() != '')]
print(f"  Registros duplicados: {len(dupl)}")

print("\n[9] AMOSTRA — DESATIVADOS:")
desativ = df[df['Data Desativacao'].str.strip() != ''][
    ['Razao Social','Data Desativacao','Motivo Desativacao']
]
print(desativ.to_string())

print("\n" + "=" * 60)
print("RESUMO FINAL")
print("=" * 60)
print(f"  Total estabelecimentos: {len(df)}")
print(f"  Ativos: {sem_data}")
print(f"  Desativados: {com_data}")
print(f"  Município: {df['Municipio'].unique()}")
print(f"  UF: {df['UF'].unique()}")
print("=" * 60)
