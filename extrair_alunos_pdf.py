import pdfplumber
import pandas as pd
import re

# Caminho do arquivo PDF
pdf_path = "./export/primeiras_paginas_unificadas.pdf"
saida_excel = "./export/alunos_extraidos.xlsx"

# Lista para guardar os dados extraídos
dados_alunos = []

# Regex para capturar: Escola, Aluno, Nome, Motivo, Data Inc
linha_regex = re.compile(r'(.+?)\s+(\d{7}-\d)\s+(.+?)\s+(ALU\.INFR)\s+(\d{2}/\d{2}/\d{4})')

with pdfplumber.open(pdf_path) as pdf:
    for i, pagina in enumerate(pdf.pages):
        texto = pagina.extract_text()
        if texto:
            print(f"📄 Página {i+1} lida com sucesso.")
            for linha in texto.split("\n"):
                match = linha_regex.match(linha)
                if match:
                    escola = match.group(1).strip()
                    aluno = match.group(2).strip().replace('-', '')
                    nome = match.group(3).strip()
                    motivo = match.group(4).strip()
                    data_inc = match.group(5).strip()
                    dados_alunos.append({
                        "Escola": escola,
                        "Aluno": aluno,
                        "Nome": nome,
                        "Motivo": motivo,
                        "Data Inc": data_inc
                    })
                # debug opcional
                else:
                    print("❌ Não casou:", linha)
        else:
            print(f"⚠️ Nada extraído da página {i+1}!")

# Exporta o Excel
df = pd.DataFrame(dados_alunos)
df.to_excel(saida_excel, index=False)

print(f"\n✅ Arquivo Excel salvo como: {saida_excel}")
print(f"🔍 Total de registros extraídos: {len(df)}")
