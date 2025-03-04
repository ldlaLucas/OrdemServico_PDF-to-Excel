import pandas as pd

def save_to_excel(data, output_file="output.xlsx"):
    """Salva os dados extraídos em um arquivo Excel."""
    df = pd.DataFrame(data, columns=["Nº da OS", "Setor", "Aberta em", "Fechada em", "Observação"])
    df.to_excel(output_file, index=False)
