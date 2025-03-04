# 📄 Extrator de Ordens de Serviço (OS)

Este é um aplicativo desenvolvido para extrair informações específicas de arquivos PDF de ordens de serviço e exportá-las para uma planilha Excel.

## 🚀 Funcionalidades

- 📄 **Leitura automática de PDFs** (todas as páginas).
- 🔍 **Extração de dados**: Nº da OS, Setor, Data de Abertura e Fechamento, e Observação.
- 📊 **Exportação para Excel**.
- 🖱️ **Interface gráfica simples e intuitiva**.
- 📂 **Escolha do arquivo de entrada e nome do arquivo de saída**.

## 🛠️ Como Usar

1. **Executar o aplicativo**

   - Se estiver usando o código Python:
     ```sh
     python main.py
     ```
   - Se estiver usando o executável (`.exe`):
     - Basta abrir o `ExtratorOS.exe`.

2. **Selecionar o PDF**

   - Clique no botão "Selecionar PDF".
   - Escolha o arquivo que deseja processar.

3. **Salvar o arquivo Excel**

   - Escolha o local e nome do arquivo para salvar os dados extraídos.

## 🏗️ Estrutura do Projeto

```
📂 ExtratorOS/
├── 📂 app/                     # Código-fonte principal
│   ├── main.py                 # Interface gráfica (Tkinter)
│   ├── extractor.py             # Extração dos dados do PDF
│   ├── file_handler.py          # Função para salvar em Excel
│   ├── requirements.txt         # Dependências do projeto
│   ├── README.md                # Documentação
├── 📂 dist/                     # Arquivos gerados pelo PyInstaller
├── 📂 build/                    # Arquivos temporários da build
└── .gitignore                   # Arquivos ignorados pelo Git
```

## 🛠️ Instalação e Configuração

### 1️⃣ Clonar o Repositório

```sh
git clone https://github.com/ldlaLucas/Analise_Ordens_de_Servi-o.git
cd ExtratorOS
```

### 2️⃣ Criar e Ativar um Ambiente Virtual

```sh
python -m venv venv
# No Windows
venv\Scripts\activate
# No Linux/Mac
source venv/bin/activate
```

### 3️⃣ Instalar Dependências

```sh
pip install -r requirements.txt
```

### 4️⃣ Dependências Necessárias

Certifique-se de ter as seguintes bibliotecas instaladas para edição e atualização do projeto:

```sh
pip install tkinter
pip install pymupdf
pip install pandas
pip install openpyxl
pip install pyinstaller
```

### 5️⃣ Gerar o Executável (Opcional)

Se quiser criar um `.exe`:

```sh
pyinstaller --onefile --windowed --name "ExtratorOS" main.py
```

## 📌 Tecnologias Utilizadas

- 🐍 Python 3
- 🖼️ Tkinter (Interface Gráfica)
- 📄 PyMuPDF (Leitura de PDFs)
- 📊 Pandas e OpenPyXL (Manipulação de Excel)
- 📦 PyInstaller (Criação do executável)

## 📢 Contribuições

Sinta-se à vontade para abrir issues e pull requests! 😃

---

✍️ **Desenvolvido por Lucas Augusto**

