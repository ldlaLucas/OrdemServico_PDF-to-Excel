from utils.extractor import extrair_dados_pdf, salvar_em_excel
import os

# Caminho do PDF dentro da pasta 'pdfs'
caminho_pdf = "pdfs/exemplo.pdf"  # Substitua pelo nome do arquivo

if os.path.exists(caminho_pdf):
    print("📄 Extraindo dados do PDF...")
    dados = extrair_dados_pdf(caminho_pdf)
    salvar_em_excel(dados)
else:
    print(f"⚠️ Arquivo {caminho_pdf} não encontrado! Coloque um PDF na pasta 'pdfs'.")