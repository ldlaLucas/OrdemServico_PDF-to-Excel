import os
import PyPDF2
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_text_from_pdf(pdf_path):
    """ Extrai o texto de todas as páginas do PDF. """
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        full_text = ""

        # Percorre todas as páginas
        for page in reader.pages:
            full_text += page.extract_text() + "\n"  # Adiciona quebra de linha entre páginas

    return full_text

def extract_fields(text):
    """ Extrai os campos desejados do texto do PDF. """
    os_number = re.search(r"N° da OS:\s*(\d+)", text)
    setor = re.search(r"Setor:\s*([\w\s]+)", text)
    aberta = re.search(r"Aberta em\s*(\d{2}/\d{2}/\d{4} \d{2}:\d{2})", text)
    fechada = re.search(r"Fechada em\s*(\d{2}/\d{2}/\d{4} \d{2}:\d{2})", text)

    # Captura "Observação" até encontrar a próxima seção (ex: "TÉCNICO RESPONSÁVEL")
    observacao_match = re.search(r"Observação:(.*?)(?=\n(?:TÉCNICO RESPONSÁVEL|Nome:|$))", text, re.S)
    observacao = observacao_match.group(1).strip() if observacao_match else "Não encontrado"

    return {
        "N° da OS": os_number.group(1) if os_number else "Não encontrado",
        "Setor": setor.group(1).strip() if setor else "Não encontrado",
        "Aberta em": aberta.group(1) if aberta else "Não encontrado",
        "Fechada em": fechada.group(1) if fechada else "Não encontrado",
        "Observação": observacao
    }

def process_pdf(pdf_path):
    """ Processa o PDF e salva os dados em um arquivo Excel. """
    text = extract_text_from_pdf(pdf_path)
    fields = extract_fields(text)

    df = pd.DataFrame([fields])
    
    # Pergunta onde salvar a planilha
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Sucesso", f"Planilha salva com sucesso!\n{save_path}")
    else:
        messagebox.showwarning("Cancelado", "Operação cancelada pelo usuário.")

def select_pdf():
    """ Abre uma janela para selecionar um PDF. """
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_path:
        process_pdf(pdf_path)
    else:
        messagebox.showwarning("Cancelado", "Nenhum arquivo selecionado.")

def create_gui():
    """ Cria a interface gráfica para arrastar e soltar o PDF. """
    root = tk.Tk()
    root.title("Extrator de Dados de PDF")

    label = tk.Label(root, text="Arraste e solte um PDF aqui ou clique para selecionar", pady=20)
    label.pack()

    btn_select = tk.Button(root, text="Selecionar PDF", command=select_pdf)
    btn_select.pack()

    root.mainloop()

if __name__ == "__main__":
    create_gui()


