import pdfplumber
import re
import pandas as pd

# Caminho do PDF
caminho_pdf = "./import_pdf/encaminhamentos.pdf"

# Lista para armazenar os dados de cada aluno
dados_alunos = []

with pdfplumber.open(caminho_pdf) as pdf:
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if texto:
            # Inicializa o dicionário para armazenar os dados do aluno
            dados = {}

            # Expressões regulares para extrair as informações
            escola_info = re.search(r'ESCOLA:\s*\d+\s*-\s*(.*?)\n', texto)
            aluno =  re.search(r'ALUNO:\s*(\d+)\s*-\s*(\d+)\s+(.*)', texto)
            filiacao = re.findall(r'FILIAÇÃO:\s*(.*?)\n(.*?)\n', texto)
            responsavel_match = re.search(r'RESPONSÁVEL:\s*(.*?)\n', texto)
            parentesco_match = re.search(r'GRAU PARENTESCO:\s*(.*?)\n', texto)
            endereco_match = re.search(r'GRAU PARENTESCO:.*?\nENDEREÇO:\s*(.*?)\n', texto, re.DOTALL)
            bairro_match = re.search(r'BAIRRO:\s*(.*?)\n', texto)
            fone_match = re.search(r'BAIRRO:.*?\nFONE:\s*(.*?)\n', texto, re.DOTALL)
            serie_completa = re.search(r'SÉRIE\s+(.*?)\n', texto)
            turma_match = re.search(r'TURMA:\s*(\w+)', texto)
            motivo_match = re.search(r'MOTIVO DO ENCAMINHAMENTO\s*\n+\s*(.*?)\n{2,}', texto, re.DOTALL)
            professor = re.search(r'PROFESSOR\s+(.*?)\n', texto)

            # Adiciona os dados extraídos ao dicionário
            if escola_info:
                dados['Escola'] = escola_info.group(1).strip()
            if aluno:
                cod_num = aluno.group(1)
                cod_dig = aluno.group(2)
                nome_final = aluno.group(3).strip()
                codigo_final = f"{cod_num}{cod_dig}"
            else:
                codigo_final = ""
                nome_final = ""
            dados['Matrícula'] = codigo_final
            dados['Aluno'] = nome_final
            if professor:
                dados['Professor'] = professor.group(1).strip()
            if serie_completa:
                serie_texto = serie_completa.group(1).strip()
                nivel_match = re.search(r'(\d+\s+\w+)$', serie_texto)
                if nivel_match:
                    dados['Etapa'] = serie_texto.replace(nivel_match.group(1), '').strip()
                    dados['Nível/Fase'] = nivel_match.group(1).strip()
                else:
                    dados['Etapa'] = serie_texto
                    dados['Nível/Fase'] = ''
            if filiacao:
                dados['Nome da Mãe'] = filiacao[0][0].strip()
                dados['Nome do Pai'] = filiacao[0][1].strip()
            if responsavel_match:
                dados['Responsável'] = responsavel_match.group(1).strip()
            if parentesco_match:
                dados['Parentesco'] = parentesco_match.group(1).strip()
            if endereco_match:
                dados['Endereço'] = endereco_match.group(1).strip()
            if bairro_match:
                dados['Bairro'] = bairro_match.group(1).strip()
            if fone_match:
                dados['Telefone'] = fone_match.group(1).strip()
            if turma_match:
                dados['Turma'] = turma_match.group(1).strip()
            if motivo_match:
                dados['Motivo do Encaminhamento'] = motivo_match.group(1).strip()

            # Adiciona o dicionário à lista se contiver dados
            if dados:
                dados_alunos.append(dados)

# Cria um DataFrame com os dados
df = pd.DataFrame(dados_alunos)

# Salva os dados em um arquivo Excel
df.to_excel("./export/dados_encaminhamentos_alunos.xlsx", index=False)

print("Arquivo 'dados_alunos.xlsx' criado com sucesso!")
