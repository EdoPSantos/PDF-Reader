import fitz  # PyMuPDF
import os
import json

# Pasta de origem dos PDFs e destino dos JSONs
input_folder = "test_pdfs"
output_folder = "jsons"
os.makedirs(output_folder, exist_ok=True)

def extract_words_with_rotation(page):
    """Extrai palavras com coordenadas e rotação."""
    word_list = []
    blocks = page.get_text("dict")["blocks"]
    for block in blocks:
        if block.get("type") != 0:
            continue  # ignora imagens
        for line in block["lines"]:
            for span in line["spans"]:
                angle = span.get("angle", 0)
                if abs(angle) > 5:
                    continue  # ignora palavras rotacionadas (verticais/inclinadas)
                word_list.append({
                    "text": span["text"],
                    "x": span["bbox"][0],
                    "y": span["bbox"][1],
                    "width": span["bbox"][2] - span["bbox"][0],
                    "height": span["bbox"][3] - span["bbox"][1],
                    "font": span.get("font"),
                    "size": span.get("size"),
                    "rotation": angle
                })
    return word_list

def process_pdf(pdf_path, output_path):
    doc = fitz.open(pdf_path)
    pages = []
    for page_number, page in enumerate(doc, start=1):
        words = extract_words_with_rotation(page)
        pages.append({
            "page": page_number,
            "words": words
        })

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(pages, f, indent=2, ensure_ascii=False)
    print(f"✅ Exportado: {output_path}")

# Processa todos os PDFs na pasta
for filename in os.listdir(input_folder):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(input_folder, filename)
    json_name = os.path.splitext(filename)[0] + ".json"
    json_path = os.path.join(output_folder, json_name)

    process_pdf(pdf_path, json_path)
