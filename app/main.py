import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from extractor import extrair_dados_pdf
from file_handler import salvar_dados_excel
import os
import threading

# Criando a janela principal
root = tk.Tk()
root.title("Extrator de Ordens de Serviço")
root.geometry("500x300")
root.resizable(False, False)
root.configure(bg="#f4f4f4")

# Estilização dos botões
btn_style = {
    "font": ("Arial", 12, "bold"),
    "bg": "#0078D7",
    "fg": "white",
    "activebackground": "#005A9E",
    "bd": 2,
    "relief": "raised"
}

# Função para selecionar e processar o PDF
def selecionar_arquivo():
    arquivo_pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    
    if not arquivo_pdf:
        return

    status_label.config(text="Processando PDF...", fg="blue")
    progress_bar.start(10)

    def processar():
        try:
            dados_extraidos = extrair_dados_pdf(arquivo_pdf)

            if not dados_extraidos:
                messagebox.showerror("Erro", "Nenhum dado foi extraído do PDF.")
                status_label.config(text="Erro ao processar!", fg="red")
                return
            
            salvar_como(dados_extraidos)

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            status_label.config(text="Erro ao processar!", fg="red")

        finally:
            progress_bar.stop()
    
    threading.Thread(target=processar).start()

# Função para salvar o arquivo Excel
def salvar_como(dados):
    arquivo_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        title="Salvar Arquivo",
        initialfile="dados_extraidos.xlsx"
    )

    if arquivo_saida:
        salvar_dados_excel(dados, arquivo_saida)
        messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{arquivo_saida}")
        status_label.config(text="Processo concluído!", fg="green")

# Criando elementos da interface
lbl_titulo = tk.Label(root, text="Extrator de Ordens de Serviço", font=("Arial", 16, "bold"), bg="#f4f4f4")
lbl_titulo.pack(pady=10)

btn_selecionar = tk.Button(root, text="Selecionar PDF", command=selecionar_arquivo, **btn_style)
btn_selecionar.pack(pady=10)

progress_bar = ttk.Progressbar(root, mode="indeterminate", length=300)
progress_bar.pack(pady=5)

status_label = tk.Label(root, text="", font=("Arial", 10), bg="#f4f4f4")
status_label.pack()

lbl_credito = tk.Label(root, text="Desenvolvido por Lucas Augusto", fg="gray", bg="#f4f4f4", font=("Arial", 9))
lbl_credito.pack(side="bottom", pady=10)

# Iniciando a interface gráfica
root.mainloop()
