import pandas as pd

# Caminhos dos arquivos
csv_path = './import_csv/PSICOTXT_2156.csv'
xlsx_path = './import_xlsx/alunos_geral.xlsx'
saida_path = './export/alunos_faltantes.xlsx'

# Lê o CSV vindo do SIGEAM
df_csv = pd.read_csv(csv_path, encoding='latin1', sep=';', on_bad_lines='skip')

# Lê o Excel com os alunos que você já tem
df_excel = pd.read_excel(xlsx_path)

# Normaliza os códigos para facilitar a comparação
csv_codigos = df_csv['Cod Aluno'].astype(str).str.replace('-', '').str.strip()
excel_codigos = df_excel['Matrícula'].astype(str).str.replace('-', '').str.strip()

# Encontra os códigos que estão no CSV mas não no Excel
codigos_para_atender = csv_codigos[~csv_codigos.isin(excel_codigos)]

# Filtra os dados dos alunos faltantes no CSV
alunos_para_atender = df_csv[df_csv['Cod Aluno'].astype(str).str.replace('-', '').str.strip().isin(codigos_para_atender)]

# Exporta a lista de alunos que ainda faltam atender
alunos_para_atender.to_excel(saida_path, index=False)

print(f'Alunos faltantes salvos em: {saida_path}')
