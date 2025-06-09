import os
import pandas as pd
import re
import string

# Palavras-chave e filtros
IGNORED_PHRASES = [
    "emitido por", "processado por computador", "condições gerais",
    "assinatura", "observações", "email", "np", "fatura", "versão", "documento"
]
REFERENCE_KEYWORDS = ["ref", "referência", "referencia", "r"]
QUANTITY_KEYWORDS = ["quantity", "quantidade", "qtd", "qt", "q"]

def normalize(text):
    return str(text).strip().lower()

def clean_label(text):
    return normalize(text).strip(string.punctuation)

def is_garbage(text):
    norm = normalize(text)
    return any(p in norm for p in IGNORED_PHRASES) or len(norm) < 3

def extract_quantity(text):
    match = re.search(r"\b(\d+[.,]?\d*)\s*(un|um|pcs|kg|mm)?\b", str(text).lower())
    return f"{match.group(1)} {match.group(2).upper() if match.group(2) else ''}".strip() if match else ""

def extract_reference(text):
    norm = normalize(text)
    if norm in REFERENCE_KEYWORDS:
        return ""
    if any(k in norm for k in REFERENCE_KEYWORDS):
        match = re.search(r"(?:ref(?:er[êé]ncia)?[:\-]?)\s*([A-Z0-9.\-/]+)", norm, re.IGNORECASE)
        if match:
            return match.group(1)
    match = re.search(r"\b([A-Z]{2,}[.\-]?[A-Z0-9]+)\b", str(text))
    return match.group(1) if match else ""

def extract_items_from_excel(df):
    items = []
    n_rows, n_cols = df.shape

    for row in range(n_rows):
        for col in range(n_cols):
            cell = str(df.iat[row, col])
            if is_garbage(cell):
                continue

            # Procurar se esta célula é uma quantidade válida
            qty_text = extract_quantity(cell)
            if qty_text and re.search(r"\d", qty_text):
                qty_row = row
                ref_text = ""
                ref_row = None

                # Procurar referência nas linhas acima (até 5 linhas antes), colunas ±1
                for dy in range(1, 6):
                    r = row - dy
                    if r < 0:
                        break
                    for dx in [-1, 0, 1]:
                        c = col + dx
                        if c < 0 or c >= n_cols:
                            continue
                        candidate_ref = extract_reference(df.iat[r, c])
                        if candidate_ref:
                            ref_text = candidate_ref
                            ref_row = r
                            break
                    if ref_row is not None:
                        break

                # Definir limites para juntar info
                y1 = ref_row if ref_row is not None else row
                y2 = qty_row
                min_y = min(y1, y2)
                max_y = max(y1, y2)

                # Junta todas as palavras das linhas entre referência e quantidade (inclusive)
                info_words = []
                for i in range(min_y, max_y + 1):
                    for word in df.iloc[i]:
                        if pd.isna(word):
                            continue
                        if not is_garbage(word):
                            info_words.append(str(word))

                items.append({
                    "referencia": ref_text,
                    "quantidade": qty_text,
                    "informacao": " ".join(info_words)
                })

    return items

# Caminho dos ficheiros Excel
excel_folder = "excels_visual"
# Ignorar ficheiros temporários do Excel (~$)
excel_files = [f for f in os.listdir(excel_folder) if f.endswith(".xlsx") and not f.startswith("~$")]

if not excel_files:
    print("Nenhum ficheiro Excel encontrado.")
    exit()

for excel_file in excel_files:
    path = os.path.join(excel_folder, excel_file)
    df = pd.read_excel(path, header=None)
    items = extract_items_from_excel(df)

    print(f"\n======= {excel_file} =======")
    if not items:
        print("Nenhum item válido encontrado.")
        continue

    for idx, item in enumerate(items, 1):
        print(f"\nItem {idx}:")
        print(" Referência:", item["referencia"])
        print(" Quantidade:", item["quantidade"])
        print(" Informação:", item["informacao"])

print("\nFim do processamento.")
