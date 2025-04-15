import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
import os
import re

# Caminho do tesseract no Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pasta = "./prints"
dados = []
linhas_nao_casadas = []

# Regex com escola
regex_com_escola = re.compile(
    r'(.+?)\s+(\d{5,8}[\/\-][\dO]+)\s+(.+?)\s+(ALU\.INFR|DIF\.APRE|DIF\.COMP|DIF\.FONO|IND\.TDAH|VULN\.SOC|NEG|TDAH C\/L)\s*(\d{2}[/|]\d{2}[/|]\d{4})?'
)


# Regex sem escola
regex_sem_escola = re.compile(
    r'(\d{5,8}[\/\-][\dO]+)\s+(.+?)\s+(ALU\.INFR|DIF\.APRE|DIF\.COMP|DIF\.FONO|IND\.TDAH|VULN\.SOC|NEG|TDAH C\/L)\s*(\d{2}[/|]\d{2}[/|]\d{4})?'
)


print("🔍 Iniciando extração de dados das imagens...")

for arquivo in sorted(os.listdir(pasta)):
    if arquivo.endswith(".png"):
        caminho = os.path.join(pasta, arquivo)
        imagem = Image.open(caminho)
        texto = pytesseract.image_to_string(imagem, lang="por")

        for linha in texto.split("\n"):
            linha = linha.strip()
            if not linha:
                continue
            
            # 🔄 Pré-tratamento para corrigir problemas comuns do OCR
            linha = re.sub(r'dif\s*[\.]?\s*apre', 'DIF.APRE', linha, flags=re.IGNORECASE)
            linha = re.sub(r'dif\s*[\.]?\s*comp', 'DIF.COMP', linha, flags=re.IGNORECASE)
            linha = re.sub(r'dif\s*[\.]?\s*fono', 'DIF.FONO', linha, flags=re.IGNORECASE)
            linha = re.sub(r'ind\s*[\.]?\s*tdah', 'IND.TDAH', linha, flags=re.IGNORECASE)
            linha = re.sub(r'alu\s*[\.]?\s*infr', 'ALU.INFR', linha, flags=re.IGNORECASE)
            linha = re.sub(r'vuln\s*[\.]?\s*soc', 'VULN.SOC', linha, flags=re.IGNORECASE)
            linha = re.sub(r'tdah\s*c\s*/\s*l', 'TDAH C/L', linha, flags=re.IGNORECASE)
            linha = re.sub(r'\s{2,}', ' ', linha)  # Remove espaços duplicados
            linha = re.sub(r'([a-z])', lambda m: m.group().upper(), linha)  # Força tudo para maiúsculo

            # Correções de data:
            linha = re.sub(r'(\d{2})[|I](\d{2})[|I](\d{4})', r'\1/\2/\3', linha)
            linha = re.sub(r'(\d{2})\s*/\s*(\d{2})\s*/\s*(\d{4})', r'\1/\2/\3', linha)
            linha = re.sub(r'(\d{2})\s+(\d{2})\s+(\d{4})', r'\1/\2/\3', linha)

            linha = re.sub(r'\s{2,}', ' ', linha)
            linha = re.sub(r'([a-z])', lambda m: m.group().upper(), linha)

            # Corrige RA que vem colado com letras (ex: 3127335-1L → 3127335-1)
            linha = re.sub(r'(\d{6,8}-\d)[A-Z]', r'\1', linha)
            linha = re.sub(r'(\d{2,5})/(\d{2,5}-\d)', r'\1\2', linha)

            match_com = regex_com_escola.match(linha)
            match_sem = regex_sem_escola.match(linha)

            if match_com:
                escola, ra, nome, motivo, data = match_com.groups()
                ra_limpo = re.sub(r'[\/]', '', ra.replace(' ', '').replace('-', '')).replace('O', '0').strip()
                dados.append({
                    "Escola": escola.strip(),
                    "Aluno": ra_limpo,
                    "Nome": nome.strip(),
                    "Motivo": motivo.strip(),
                    "Data Inc": data.replace('|', '/').strip() if data else "",
                    "Imagem": arquivo
                })
            elif match_sem:
                ra, nome, motivo, data = match_sem.groups()
                ra_limpo = re.sub(r'[\/]', '', ra.replace(' ', '').replace('-', '')).replace('O', '0').strip()
                dados.append({
                    "Escola": "",
                    "Aluno": ra_limpo,
                    "Nome": nome.strip(),
                    "Motivo": motivo.strip(),
                    "Data Inc": data.replace('|', '/').strip() if data else "",
                    "Imagem": arquivo
                })
            else:
                # Tentativa de capturar RA, Nome e Motivo (sem data)
                match_parcial = re.match(r'(.+?)\s+(\d{6,8}-\d)\s+(.+?)\s+([A-ZÀ-Ú/\.\s]{3,20})$', linha)
                if match_parcial:
                    escola, ra, nome, motivo = match_parcial.groups()
                    dados.append({
                        "Escola": escola.strip(),
                        "Aluno": ra.replace(' ', '').replace('-', '').replace('/', '').replace('O', '0').strip(),
                        "Nome": nome.strip(),
                        "Motivo": motivo.strip(),
                        "Data Inc": "",  # Data ausente
                        "Imagem": arquivo
                    })
                else:
                    # 📌 NOVO: captura apenas RA isolado
                    match_ra_simples = re.match(r'^(\d{6,8}-\d)$', linha)
                    if match_ra_simples:
                        ra = match_ra_simples.group(1)
                        dados.append({
                            "Escola": "",
                            "Aluno": ra.strip(),
                            "Nome": "",
                            "Motivo": "",
                            "Data Inc": "",
                            "Imagem": arquivo
                        })
                    else:
                        linhas_nao_casadas.append(f"[{arquivo}] {linha}")

# Salva Excel com dados extraídos
df = pd.DataFrame(dados)
os.makedirs("./export", exist_ok=True)
df.to_excel("./export/dados_extraidos_dos_prints.xlsx", index=False)
print(f"✅ Total de registros extraídos: {len(dados)}")
print("📁 Excel salvo: ./export/dados_extraidos_dos_prints.xlsx")

# Salva texto das linhas que não foram casadas
with open("./export/linhas_ignoradas.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(linhas_nao_casadas))

print(f"📉 Total de linhas ignoradas: {len(linhas_nao_casadas)}")
print("📁 Verifique: ./export/linhas_ignoradas.txt")