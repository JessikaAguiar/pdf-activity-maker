import os
from PyPDF2 import PdfReader, PdfWriter

# Pasta onde estão os arquivos PDF
pasta_pdfs = "./pdfs_unificar"  # <- coloque seus arquivos PDF nessa pasta
saida_pdf = "./export/primeiras_paginas_unificadas.pdf"

# Cria um objeto para escrever o novo PDF
pdf_writer = PdfWriter()

# Percorre todos os arquivos .pdf na pasta
for nome_arquivo in os.listdir(pasta_pdfs):
    if nome_arquivo.lower().endswith(".pdf"):
        caminho_pdf = os.path.join(pasta_pdfs, nome_arquivo)
        reader = PdfReader(caminho_pdf)

        # Verifica se tem pelo menos uma página
        if len(reader.pages) > 0:
            primeira_pagina = reader.pages[0]
            pdf_writer.add_page(primeira_pagina)
            print(f"Adicionada a 1ª página de: {nome_arquivo}")
        else:
            print(f"⚠️ PDF sem páginas: {nome_arquivo}")

# Salva o novo PDF com todas as primeiras páginas
with open(saida_pdf, "wb") as f:
    pdf_writer.write(f)

print(f"\n✅ PDF unificado salvo como: {saida_pdf}")
