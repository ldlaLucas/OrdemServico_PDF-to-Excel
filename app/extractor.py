import pdfplumber
import re

def extrair_dados_pdf(caminho_pdf):
    dados_extraidos = []
    
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue
            
            linhas = texto.split("\n")
            registro_atual = {}

            for linha in linhas:
                # Para a leitura ao encontrar "TÉCNICO RESPONSÁVEL"
                if "TÉCNICO RESPONSÁVEL" in linha:
                    break

                # Detecta Nº da OS (começa um novo registro)
                match_os = re.search(r"Nº da OS:\s*(\d{9})", linha)
                if match_os:
                    if registro_atual:
                        dados_extraidos.append(registro_atual)
                    registro_atual = {
                        "Nº da OS": match_os.group(1),
                        "Setor": "",
                        "Aberta em": "",
                        "Fechada em": "",
                        "Observação": ""
                    }
                
                # Captura Setor
                match_setor = re.search(r"Setor:\s*(.+)", linha)
                if match_setor:
                    registro_atual["Setor"] = match_setor.group(1).strip()
                
                # Captura Datas (Melhoria na detecção)
                match_aberta = re.search(r"Aberta em\s*([\d/]+\s[\d:]+)", linha)
                match_fechada = re.search(r"Fechada em\s*([\d/]+\s[\d:]+)", linha)

                if match_aberta:
                    registro_atual["Aberta em"] = match_aberta.group(1)
                
                if match_fechada:
                    registro_atual["Fechada em"] = match_fechada.group(1)

                # Captura Observação (melhorando extração multilinha)
                if "Observação:" in linha:
                    registro_atual["Observação"] = linha.split("Observação:")[-1].strip()
                elif registro_atual.get("Observação"):  # Continua extraindo caso seja multilinha
                    registro_atual["Observação"] += " " + linha.strip()

            # Adiciona o último registro, se houver
            if registro_atual:
                dados_extraidos.append(registro_atual)

    return dados_extraidos
