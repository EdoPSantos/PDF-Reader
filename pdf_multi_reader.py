import os
import pdfplumber
import re
import logging

logging.getLogger("pdfminer").setLevel(logging.ERROR)

IGNORED_PHRASES = [
    "emitido por", "processado por computador", "condições gerais",
    "assinatura", "observações", "email", "np", "fatura", "versão", "documento"
]

ITEM_KEYWORDS = [
    "ref", "marca", "fornecedor", "iso", "un", "pcs", "kg", "mm",
    "descrição", "designação", "especificação", "qtd", "quantidade", "entrega",
    "código", "item", "material", "data"
]

REQUIRED_KEYWORDS = ["ref", "marca", "quant", "un"]

def normalize(text):
    return text.strip().lower()

def is_garbage(line):
    norm = normalize(line)
    return any(p in norm for p in IGNORED_PHRASES) or len(norm) < 3

def looks_like_item_start(line):
    norm = normalize(line)
    has_qty = bool(re.search(r"\b\d+[,.]?\d*\s*(un|kg|pcs|mm)\b", norm))
    has_code = bool(re.search(r"\b\d{5,}\b", norm))
    has_date = bool(re.search(r"\b\d{4}-\d{2}-\d{2}\b", norm))
    has_words = len(norm.split()) > 2
    return (has_qty or has_code or has_date) and has_words

def is_valid_item_block(block):
    joined = " ".join(block).lower()
    found = sum(1 for key in REQUIRED_KEYWORDS if key in joined)
    return found >= 2

def extract_items(lines):
    items = []
    buffer = []
    inside_item = False

    for i, line in enumerate(lines):
        norm = normalize(line)
        if is_garbage(norm):
            continue

        if looks_like_item_start(norm):
            if buffer:
                if is_valid_item_block(buffer):
                    items.append(buffer.copy())
                buffer.clear()
            buffer.append(line)
            inside_item = True
        elif inside_item:
            if re.match(r"^\s*$", norm):
                continue
            if any(k in norm for k in ITEM_KEYWORDS) or len(norm.split()) < 8:
                buffer.append(line)
            else:
                if is_valid_item_block(buffer):
                    items.append(buffer.copy())
                buffer.clear()
                inside_item = False

    if buffer and is_valid_item_block(buffer):
        items.append(buffer.copy())

    return items

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
            text = page.extract_text()
            if text:
                page_lines = [line.strip() for line in text.split("\n") if line.strip()]
                all_lines.extend(page_lines)

    items = extract_items(all_lines)

    if not items:
        print("Nenhum item detectado.")
        continue

    for i, item_lines in enumerate(items, 1):
        print(f"\nItem {i}:")
        for line in item_lines:
            print(" ", line)

print("\nFim do processamento.")
