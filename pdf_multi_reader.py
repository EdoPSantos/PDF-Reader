import os
import pdfplumber
import re
import logging

logging.getLogger("pdfminer").setLevel(logging.ERROR)

IGNORED_PHRASES = [
    "emitido por", "processado por computador", "condições gerais",
    "assinatura", "observações", "email", "np", "fatura", "versão", "documento"
]


def normalize(text):
    return text.strip().lower()

def is_garbage(line):
    norm = normalize(line)
    return any(p in norm for p in IGNORED_PHRASES) or len(norm) < 3

def extract_lines_by_position(page, y_tolerance=3):
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    if not words:
        return []
    # Agrupar palavras por linha (top)
    lines_dict = {}
    for word in words:
        top = round(word["top"] / y_tolerance) * y_tolerance
        lines_dict.setdefault(top, []).append(word)
    # Ordenar linhas e palavras por x
    lines = []
    for top in sorted(lines_dict):
        line_words = sorted(lines_dict[top], key=lambda w: w["top"])
        line_text = " ".join(w["text"] for w in line_words)
        lines.append(line_text)
    return lines

def extract_reference(text):
    match = re.search(r"\b([A-Z]{2,}[\.-]?[A-Z0-9]+)\b", text)
    return match.group(1) if match else ""

def extract_quantity(text):
    # Só extrai se vier imediatamente seguido de unidade (ex: 10 UN, 5 UM, 2 KG, etc.)
    match = re.search(r"\b(\d+[,.]?\d*)\s*(un|um|pcs|kg|mm)\b", text.lower())
    return f"{match.group(1)} {match.group(2).upper()}" if match else ""

def extract_items(lines):
    items = []
    buffer = []

    for line in lines:
        norm = normalize(line)
        if is_garbage(norm):
            continue

        # Novo item se linha tem muitos dígitos ou unidade
        if re.search(r"\b\d{4,}\b", norm) or re.search(r"\b\d+[,.]?\d*\s*(un|kg|pcs|mm)\b", norm):
            if buffer:
                items.append(buffer.copy())
                buffer.clear()
            buffer.append(line)
        else:
            buffer.append(line)
    if buffer:
        items.append(buffer.copy())
    return items

def parse_item_block(block):
    full_text = " ".join(block)
    ref = extract_reference(full_text)
    qty = extract_quantity(full_text)
    info = full_text
    return {"referencia": ref, "quantidade": qty, "informacao": info}

folder = "test_pdfs"
pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]

if not pdf_files:
    print("Nenhum PDF encontrado.")
    exit()

for pdf_file in pdf_files:
    path = os.path.join(folder, pdf_file)
    print(f"\n==============================\nPDF: {pdf_file}\n==============================")

    all_lines = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_lines = extract_lines_by_position(page)
            all_lines.extend(page_lines)

    item_blocks = extract_items(all_lines)

    if not item_blocks or all(len(block) == 0 for block in item_blocks):
        print("Nenhum item detectado.")
        continue

    for i, block in enumerate(item_blocks, 1):
        parsed = parse_item_block(block)
        print(f"\nItem {i}:")
        print(" Referência:", parsed["referencia"])
        print(" Quantidade:", parsed["quantidade"])
        print(" Informação:", parsed["informacao"])

print("\nFim do processamento.")
