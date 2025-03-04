import fitz  # PyMuPDF
import re

def extract_data(pdf_path):
    """Extrai os campos necessários do PDF e suas coordenadas."""
    doc = fitz.open(pdf_path)
    extracted_info = []

    for page in doc:
        text = page.get_text("text")
        
        # Expressões regulares para capturar os campos
        os_number = re.search(r"Nº da OS:\s*(\d+)", text)
        setor_match = re.search(r"Setor:\s*(.*?)\s*Oficina:", text)
        aberta_em = re.search(r"Aberta em (\d{2}/\d{2}/\d{4} \d{2}:\d{2})", text)
        fechada_em = re.search(r"Fechada em (\d{2}/\d{2}/\d{4} \d{2}:\d{2})", text)
        observacao_match = re.search(r"Observação:\s*(.*)", text, re.DOTALL)

        # Ajuste do setor para remover informações extras
        setor = setor_match.group(1).split(" Prioridade:")[0] if setor_match else ""

        # Capturar coordenadas dos campos
        coords = {}
        for field, regex in {
            "Nº da OS": os_number,
            "Setor": setor_match,
            "Aberta em": aberta_em,
            "Fechada em": fechada_em,
            "Observação": observacao_match
        }.items():
            if regex:
                rects = page.search_for(regex.group(0))
                if rects:
                    coords[field] = rects[0]

        # Salvar dados extraídos
        extracted_info.append({
            "Nº da OS": os_number.group(1) if os_number else "",
            "Setor": setor,
            "Aberta em": aberta_em.group(1) if aberta_em else "",
            "Fechada em": fechada_em.group(1) if fechada_em else "",
            "Observação": observacao_match.group(1).strip() if observacao_match else "",
            "coords": coords  # Coordenadas para destaque visual
        })

    return extracted_info
