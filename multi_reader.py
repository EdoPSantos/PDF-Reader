import os
import pandas as pd
import re
import unicodedata

REFERENCE_KEYWORDS = ["referência", "referencia", "ref", "reference", "r"]
QUANTITY_KEYWORDS = ["quantidade", "qtd", "qt", "quantity", "q"]

def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    text = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    return text

def is_garbage(text):
    norm = normalize(text)
    return norm == "" or norm == "nan"

def is_quantity_number(text):
    try:
        t = str(text).replace(",", ".").strip()
        float(re.findall(r"[-+]?\d*[\.,]?\d+", t)[0])
        return True
    except:
        return False

def find_headers(df, keywords):
    """
    Procura o cabeçalho e retorna os índices das colunas que contêm as keywords.
    """
    n_rows, n_cols = df.shape
    header_row = None
    cols_found = []
    for row in range(n_rows):
        for col in range(n_cols):
            cell = normalize(df.iat[row, col])
            for k in keywords:
                if re.search(rf'(^|\s|[\.\:;\-\)]){re.escape(k)}($|\s|[\.\:;\-\)])', cell):
                    cols_found.append(col)
                    header_row = row
        if cols_found:
            break
    return header_row, cols_found

def extract_items_from_excel_universal(df):
    n_rows, n_cols = df.shape

    # Encontrar cabeçalho de referência e quantidade
    header_row_ref, ref_cols = find_headers(df, REFERENCE_KEYWORDS)
    header_row_qty, qty_cols = find_headers(df, QUANTITY_KEYWORDS)

    if header_row_ref is None or header_row_qty is None or not ref_cols or not qty_cols:
        print("Não foi possível encontrar cabeçalho de referência ou quantidade.")
        return []

    # Usa o cabeçalho que vier mais acima (mais universal)
    header_row = min(header_row_ref, header_row_qty)

    items = []
    # Para cada linha de dados depois do cabeçalho
    for row in range(header_row + 1, n_rows):
        # Procurar referência e quantidade em colunas ±1 das respetivas colunas detetadas
        ref_val = ""
        qty_val = ""
        # Procura referência
        for ref_col in ref_cols:
            for delta in [-1, 0, 1]:
                c = ref_col + delta
                if c < 0 or c >= n_cols:
                    continue
                cell = df.iat[row, c]
                if not is_garbage(cell):
                    ref_val = str(cell).strip()
                    break
            if ref_val:
                break
        # Procura quantidade
        for qty_col in qty_cols:
            for delta in [-1, 0, 1]:
                c = qty_col + delta
                if c < 0 or c >= n_cols:
                    continue
                cell = df.iat[row, c]
                if not is_garbage(cell) and is_quantity_number(cell):
                    qty_val = str(cell).strip()
                    break
            if qty_val:
                break

        # Só guarda se encontrou ambos e quantidade é numérica
        if ref_val and qty_val and is_quantity_number(qty_val):
            info_words = [str(df.iat[row, c]) for c in range(n_cols) if not is_garbage(df.iat[row, c])]
            items.append({
                "referencia": ref_val,
                "quantidade": qty_val,
                "informacao": " ".join(info_words)
            })
    return items

# Caminho dos ficheiros Excel
excel_folder = "excels_visual"
excel_files = [f for f in os.listdir(excel_folder) if f.endswith(".xlsx") and not f.startswith("~$")]

if not excel_files:
    print("Nenhum ficheiro Excel encontrado.")
    exit()

for excel_file in excel_files:
    path = os.path.join(excel_folder, excel_file)
    df = pd.read_excel(path, header=None)
    print(f"\n======= {excel_file} =======")
    items = extract_items_from_excel_universal(df)

    if not items:
        print("Nenhum item válido encontrado.")
        continue

    for idx, item in enumerate(items, 1):
        print(f"\nItem {idx}:")
        print(" Referência:", item["referencia"])
        print(" Quantidade:", item["quantidade"])
        print(" Informação:", item["informacao"])

print("\nFim do processamento.")
