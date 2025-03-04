import tkinter as tk
from tkinter import filedialog, Canvas
import fitz  # PyMuPDF
from extractor import extract_data
from file_handler import save_to_excel
from PIL import Image, ImageTk

class PDFExtractorApp:
    def __init__(self, root):
        """Inicializa a interface do aplicativo."""
        self.root = root
        self.root.title("Extrator de Ordens de Serviço")

        # Botão para selecionar PDF
        self.btn_open = tk.Button(root, text="Selecionar PDF", command=self.open_pdf)
        self.btn_open.pack(pady=10)

        # Canvas para exibição do PDF
        self.canvas = Canvas(root, width=600, height=800)
        self.canvas.pack()

        # Botões de navegação
        self.btn_prev = tk.Button(root, text="◀ Página Anterior", command=self.prev_page)
        self.btn_prev.pack(side=tk.LEFT, padx=10)

        self.btn_next = tk.Button(root, text="▶ Próxima Página", command=self.next_page)
        self.btn_next.pack(side=tk.RIGHT, padx=10)

        # Variáveis de controle do PDF
        self.pdf_document = None
        self.current_page = 0

    def open_pdf(self):
        """Seleciona um arquivo PDF e exibe a primeira página."""
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if file_path:
            self.pdf_document = fitz.open(file_path)
            self.current_page = 0
            self.display_page(self.current_page)

            # Extração de dados
            extracted_data = extract_data(file_path)

            # Destaque dos campos na interface
            self.highlight_fields(extracted_data)

            # Salvar no Excel
            save_to_excel(extracted_data, "dados_extraidos.xlsx")

    def display_page(self, page_number):
        """Exibe a página selecionada do PDF na interface."""
        if self.pdf_document:
            page = self.pdf_document.load_page(page_number)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.resize((600, 800))
            self.pdf_image = ImageTk.PhotoImage(img)
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_image)

    def highlight_fields(self, extracted_data):
        """Destaca os campos extraídos no PDF com um quadro vermelho."""
        for field in extracted_data:
            x, y, w, h = field['coords']
            self.canvas.create_rectangle(x, y, x + w, y + h, outline="red", width=2)

    def next_page(self):
        """Vai para a próxima página do PDF."""
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.display_page(self.current_page)

    def prev_page(self):
        """Volta para a página anterior do PDF."""
        if self.pdf_document and self.current_page > 0:
            self.current_page -= 1
            self.display_page(self.current_page)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
