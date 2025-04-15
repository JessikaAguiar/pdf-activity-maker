import pandas as pd

# Caminho do CSV original
csv_path = './import_csv/PSICOTXT.csv'

# Caminho do novo arquivo .xlsx
xlsx_path = './export/PSICOTXT_convertido.xlsx'

# Lê o CSV (com codificação correta)
df = pd.read_csv(csv_path, encoding='latin1', sep=';', on_bad_lines='skip')

# Salva como .xlsx
df.to_excel(xlsx_path, index=False)

print(f'CSV convertido com sucesso para: {xlsx_path}')