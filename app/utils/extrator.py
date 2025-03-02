import PyPDF2
import pandas as pd
import re
import os

def extrair_dados_pdf(caminho_pdf):
    with open(caminho_pdf, "rb") as file:
        leitor = PyPDF2.PdfReader(file)
        dados_extraidos = []
        
        for pagina in leitor.pages:
            texto = pagina.extract_text()
            if texto:
                # Expressões regulares para capturar os dados
                num_os = re.search(r"Nº da Os:\s*(\d+)", texto)
                setor = re.search(r"Setor:\s*([\w\s]+)", texto)
                aberta_em = re.search(r"Aberta em ([\d/\s:]+)", texto)
                fechada_em = re.search(r"Fechada em ([\d/\s:]+)", texto)
                observacao = re.search(r"Observação:\s*(.*)", texto)
                
                dados = {
                    "Nº da OS": num_os.group(1) if num_os else "",
                    "Setor": setor.group(1).strip() if setor else "",
                    "Aberta em": aberta_em.group(1).strip() if aberta_em else "",
                    "Fechada em": fechada_em.group(1).strip() if fechada_em else "",
                    "Observação": observacao.group(1).strip() if observacao else ""
                }
                dados_extraidos.append(dados)
    
    return dados_extraidos

def salvar_em_excel(dados, nome_arquivo="output/dados_extraidos.xlsx"):
    os.makedirs("output", exist_ok=True)  # Cria a pasta output se não existir
    df = pd.DataFrame(dados)
    df.to_excel(nome_arquivo, index=False)
    print(f"✅ Dados salvos em {nome_arquivo}")
