import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from extractor import extrair_dados_pdf
from file_handler import salvar_dados_excel
import fitz
from PIL import Image, ImageTk
import threading
import sys

sys.stdout.reconfigure(encoding='utf-8')


# Criando a janela principal
root = tk.Tk()
root.title("Extrator de Ordens de Serviço")
root.geometry("800x650")
root.resizable(False, False)

# Variáveis globais
pdf_documento = None
pagina_atual = 0
imagens_paginas = []
btn_extrair = None  # Inicialmente oculto

# Estilização dos botões
btn_style = {
    "font": ("Arial", 12, "bold"),
    "bg": "#0078D7",
    "fg": "white",
    "activebackground": "#005A9E",
    "bd": 2,
    "relief": "raised"
}

# Função para carregar o PDF e exibir a primeira página
def selecionar_arquivo():
    global pdf_documento, pagina_atual, imagens_paginas, btn_extrair

    arquivo_pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if not arquivo_pdf:
        return

    try:
        pdf_documento = fitz.open(arquivo_pdf)
        pagina_atual = 0
        imagens_paginas = []

        for pagina in pdf_documento:
            pix = pagina.get_pixmap()
            imagem = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            imagens_paginas.append(imagem)

        exibir_pagina()
        status_label.config(text="PDF carregado com sucesso!", fg="green")

        # Criar botão de extração após o carregamento, caso ainda não exista
        if btn_extrair is None:
            criar_botao_extrair()

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao carregar PDF: {e}")

# Função para exibir a página atual
def exibir_pagina():
    if not imagens_paginas:
        return

    imagem = imagens_paginas[pagina_atual]
    imagem.thumbnail((420, 595))  # Ajuste para exibição
    img_tk = ImageTk.PhotoImage(imagem)

    canvas.delete("all")
    canvas.create_image(0, 0, anchor="nw", image=img_tk)
    canvas.image = img_tk

    label_pagina.config(text=f"Página {pagina_atual + 1} de {len(imagens_paginas)}")

# Função para avançar página
def proxima_pagina():
    global pagina_atual
    if pdf_documento and pagina_atual < len(imagens_paginas) - 1:
        pagina_atual += 1
        exibir_pagina()

# Função para voltar página
def pagina_anterior():
    global pagina_atual
    if pdf_documento and pagina_atual > 0:
        pagina_atual -= 1
        exibir_pagina()

# Função para criar o botão "Extrair Dados"
def criar_botao_extrair():
    global btn_extrair
    btn_extrair = tk.Button(root, text="Extrair Dados", command=extrair_dados, **btn_style)
    btn_extrair.pack(pady=10)

# Função para extrair os dados do PDF
def extrair_dados():
    if not pdf_documento:
        messagebox.showwarning("Aviso", "Nenhum PDF carregado!")
        return

    status_label.config(text="Extraindo dados...", fg="blue")
    progress_bar.start(10)

    def processar():
        try:
            print("📥 Iniciando extração...")
            dados_extraidos = extrair_dados_pdf(pdf_documento)

            if not dados_extraidos:
                print("❌ Nenhum dado extraído")
                messagebox.showerror("Erro", "Nenhum dado foi extraído do PDF.")
                status_label.config(text="Erro ao extrair dados!", fg="red")
                return
            
            print(f"✅ Dados extraídos: {dados_extraidos}")
            salvar_como(dados_extraidos)

        except Exception as e:
            print(f"❌ Erro na extração: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            status_label.config(text="Erro ao extrair dados!", fg="red")

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
frame_topo = tk.Frame(root)
frame_topo.pack(fill="x", pady=5)

btn_selecionar = tk.Button(frame_topo, text="Selecionar PDF", command=selecionar_arquivo, **btn_style)
btn_selecionar.pack()
btn_extrair = tk.Button(root, text="Extrair Dados", command=extrair_dados, **btn_style)
btn_extrair.pack(pady=10)

canvas = tk.Canvas(root, width=700, height=500, bg="white")
canvas.pack(pady=10)

frame_controles = tk.Frame(root)
frame_controles.pack()

btn_anterior = tk.Button(frame_controles, text="◀ Página Anterior", command=pagina_anterior, **btn_style)
btn_anterior.grid(row=0, column=0, padx=5)

label_pagina = tk.Label(frame_controles, text="Página 0 de 0", font=("Arial", 12))
label_pagina.grid(row=0, column=1, padx=5)

btn_proxima = tk.Button(frame_controles, text="Próxima Página ▶", command=proxima_pagina, **btn_style)
btn_proxima.grid(row=0, column=2, padx=5)

progress_bar = ttk.Progressbar(root, mode="indeterminate", length=300)
progress_bar.pack(pady=5)

status_label = tk.Label(root, text="", font=("Arial", 10))
status_label.pack()

lbl_credito = tk.Label(root, text="Desenvolvido por Lucas Augusto", fg="gray", font=("Arial", 9))
lbl_credito.pack(side="bottom", pady=10)

root.mainloop()
