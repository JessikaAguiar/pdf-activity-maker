import pandas as pd

# Caminho da planilha original e do novo arquivo
arquivo_entrada = "./export/alunos_extraidos.xlsx"
arquivo_saida = "./export/alunos_sem_duplicatas.xlsx"

# Lê a planilha
df = pd.read_excel(arquivo_entrada)

# Remove duplicatas com base nas colunas-chave
df_limpo = df.drop_duplicates(subset=["Aluno"], keep='first')

# Salva em um novo arquivo Excel
df_limpo.to_excel(arquivo_saida, index=False)

print(f"✅ Arquivo salvo sem duplicatas: {arquivo_saida}")
print(f"🔍 Linhas originais: {len(df)} → Linhas após limpeza: {len(df_limpo)}")
