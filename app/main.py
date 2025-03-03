
import tkinter as tk
from tkinter import filedialog, messagebox
from extractor import extrair_dados_pdf
from file_handler import salvar_dados_excel
import os

def selecionar_arquivo():
    """Abre uma janela para selecionar o arquivo PDF."""
    arquivo_pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    
    if not arquivo_pdf:
        return  # Se o usuário não selecionar nada, não faz nada

    try:
        dados_extraidos = extrair_dados_pdf(arquivo_pdf)

        if not dados_extraidos:
            messagebox.showerror("Erro", "Nenhum dado foi extraído do PDF.")
            return
        
        salvar_como(dados_extraidos)
    
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def salvar_como(dados):
    """Permite ao usuário escolher o nome e o local do arquivo de saída."""
    arquivo_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Salvar Arquivo",
        initialfile="dados_extraidos.xlsx"
    )

    if arquivo_saida:
        salvar_dados_excel(dados, arquivo_saida)
        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{arquivo_saida}")

# Criando a interface gráfica com Tkinter
root = tk.Tk()
root.title("Extrator de Dados de PDF")

# Configurações de layout
root.geometry("400x200")
root.resizable(False, False)

# Botão para selecionar arquivo PDF
btn_selecionar = tk.Button(root, text="Selecionar PDF", command=selecionar_arquivo, height=2, width=20)
btn_selecionar.pack(pady=20)

# Rodapé
lbl_credito = tk.Label(root, text="Desenvolvido por Lucas Augusto", fg="gray")
lbl_credito.pack(side="bottom", pady=10)

# Inicia o loop da interface gráfica
root.mainloop()
