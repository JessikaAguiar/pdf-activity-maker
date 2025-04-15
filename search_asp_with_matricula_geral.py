import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

# Caminhos dos arquivos
caminho_geral = "./import_xlsx/geral_asp.xlsx"
caminho_controle = "./import_xlsx/controle.xlsx"

# Carrega os dados das planilhas
df_asp = pd.read_excel(caminho_geral)
df_controle = pd.read_excel(caminho_controle)


# Converte colunas de matrícula para string para garantir comparação correta
df_asp["Matrícula"] = df_asp["Matrícula"].astype(str)
df_controle["Matrícula"] = df_controle["Matrícula"].astype(str)

# Atualiza o status para 'ATENDIDOS' se a matrícula estiver na planilha de controle
df_asp.loc[(df_asp["Matrícula"].replace('-','')).isin(df_controle["Matrícula"]), "Status"] = "Encontrado"


# Salva novamente a planilha atualizada
df_asp.to_excel(caminho_geral, index=False)

wb = load_workbook(caminho_geral)
ws = wb.active

wb.save(caminho_geral)

print("Status atualizado com sucesso na planilha.")