import pandas as pd

def salvar_dados_excel(dados, nome_arquivo="dados_extraidos.xlsx"):
    """Salva os dados extraídos em um arquivo Excel."""
    df = pd.DataFrame(dados)  # Converte a lista de dicionários em DataFrame do Pandas
    df.to_excel(nome_arquivo, index=False, engine='openpyxl')
    print(f"Arquivo salvo como {nome_arquivo}")
