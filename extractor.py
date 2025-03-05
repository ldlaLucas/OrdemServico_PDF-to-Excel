import pdfplumber
import re

def extrair_dados_pdf(caminho_pdf):
    dados_extraidos = []

    # Abrir o PDF com pdfplumber
    with pdfplumber.open(caminho_pdf.name) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            print(f"\n--- Página {i+1} ---\n{texto}\n")  # Exibir o conteúdo extraído

            if not texto:
                print(f"⚠ Nenhum texto extraído na página {i+1}")
                continue
            
            linhas = texto.split("\n")
            registro_atual = {}

            for linha in linhas:
                print(f"📌 Linha: {linha}")  # Ver o que está sendo lido

                if "TÉCNICO RESPONSÁVEL" in linha:
                    break

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
                
                match_setor = re.search(r"Setor:\s*(.+)", linha)
                if match_setor:
                    setor = match_setor.group(1).strip()
                    # Alteração aqui: Remover texto após a palavra "Prioridade"
                    if "Prioridade" in setor:
                        setor = setor.split("Prioridade")[0].strip()
                    registro_atual["Setor"] = setor
                
                match_aberta = re.search(r"Aberta em\s*([\d/]+\s[\d:]+)", linha)
                match_fechada = re.search(r"Fechada em\s*([\d/]+\s[\d:]+)", linha)

                if match_aberta:
                    registro_atual["Aberta em"] = match_aberta.group(1)
                
                if match_fechada:
                    registro_atual["Fechada em"] = match_fechada.group(1)

                if "Observação:" in linha:
                    observacao = linha.split("Observação:")[-1].strip()
                    if "itens" in observacao:
                        observacao = observacao.split("itens")[0].strip()
                    registro_atual["Observação"] = observacao
                elif registro_atual.get("Observação"):  
                    linha_observacao = linha.strip()
                    if "itens" in linha_observacao:
                        linha_observacao = linha_observacao.split("itens")[0].strip()
                    registro_atual["Observação"] += " " + linha_observacao

            if registro_atual:
                dados_extraidos.append(registro_atual)

    print(f"✅ Dados extraídos: {dados_extraidos}")  # Verificar os dados extraídos
    return dados_extraidos
