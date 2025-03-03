import PyPDF2
import re

def extrair_dados_pdf(arquivo_pdf):
    """Lê um PDF e extrai os campos necessários de todas as páginas."""
    with open(arquivo_pdf, "rb") as pdf_file:
        leitor = PyPDF2.PdfReader(pdf_file)
        dados_extraidos = []

        # Percorre todas as páginas do PDF
        for pagina in leitor.pages:
            texto = pagina.extract_text()

            # Expressões regulares para encontrar os dados
            num_os = re.search(r"Nº da OS:\s*(\d+)", texto)
            setor = re.search(r"Setor:\s*([\w\s]+)", texto)
            aberta = re.search(r"Aberta em\s*([\d/:\s]+)", texto)
            fechada = re.search(r"Fechada em\s*([\d/:\s]+)", texto)
            observacao = re.search(r"Observação:\s*(.+)", texto)

            # Criando um dicionário com os dados extraídos
            dados = {
                "Nº da OS": num_os.group(1) if num_os else "Não encontrado",
                "Setor": setor.group(1) if setor else "Não encontrado",
                "Aberta em": aberta.group(1).strip() if aberta else "Não encontrado",
                "Fechada em": fechada.group(1).strip() if fechada else "Não encontrado",
                "Observação": observacao.group(1).strip() if observacao else "Não encontrado"
            }

            dados_extraidos.append(dados)

    return dados_extraidos
