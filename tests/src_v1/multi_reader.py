import os
import pandas as pd
import re
import unicodedata

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

def is_header_word(text):
    headers = [
        "ref", "referência", "referencia", "reference", "r",
        "código", "code", "descrição", "description",
        "qt", "qt.", "quantidade", "un", "unit", "pcs", "material", "item"
    ]
    return normalize(text) in [normalize(h) for h in headers]

def is_quantity_number(text):
    try:
        t = str(text).replace(",", ".").strip()
        float(re.findall(r"[-+]?\d*[\.,]?\d+", t)[0])
        return True
    except:
        return False

def find_headers(df, keywords):
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

def extract_items_quantity_only(df, SEARCH_DEPTH=10):
    n_rows, n_cols = df.shape
    header_row_qty, qty_cols = find_headers(df, QUANTITY_KEYWORDS)
    if header_row_qty is None or not qty_cols:
        print("Não foi possível encontrar cabeçalho de quantidade.")
        return []
    items = []
    already_found = set()  # para evitar duplicados

    for qty_col in qty_cols:
        for row in range(header_row_qty + 1, n_rows):
            for dr in range(SEARCH_DEPTH):
                r2 = row + dr
                if r2 >= n_rows:
                    break
                for dc in [-2, -1, 0, 1, 2]:
                    c2 = qty_col + dc
                    if c2 < 0 or c2 >= n_cols:
                        continue
                    cell = df.iat[r2, c2]
                    val = str(cell).strip()
                    if not is_garbage(val) and not is_header_word(val) and is_quantity_number(val):
                        key = (val, r2)
                        if key in already_found:
                            continue
                        already_found.add(key)
                        info = " ".join(str(df.iat[r2, cc]) for cc in range(n_cols) if not is_garbage(df.iat[r2, cc]))
                        try:
                            quantidade_int = int(float(val.replace(",", ".")))
                        except Exception:
                            quantidade_int = val  # mantém original se não for possível converter
                        items.append({
                            "Quantidade": quantidade_int,
                            "Informacao": info
                        })
                        break
    return items

# Caminho dos ficheiros Excel
excel_folder = "../excels_visual"
output_folder = "../result"
os.makedirs(output_folder, exist_ok=True)
excel_files = [f for f in os.listdir(excel_folder) if f.endswith(".xlsx") and not f.startswith("~$")]

if not excel_files:
    print("Nenhum ficheiro Excel encontrado.")
    exit()

for excel_file in excel_files:
    path = os.path.join(excel_folder, excel_file)
    sheets = pd.read_excel(path, header=None, sheet_name=None)
    all_items = []
    for sheet_name, df in sheets.items():
        items = extract_items_quantity_only(df)
        all_items.extend(items)

    if not all_items:
        print(f"[{excel_file}] Nenhum item válido encontrado.")
        output_path = os.path.join(output_folder, os.path.splitext(excel_file)[0] + ".txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("Nenhum valor adquirido.")
        print(f"[{excel_file}] Ficheiro criado em: {output_path}")
        continue

    # Adiciona coluna "Item" (número sequencial, começa em 1)
    for i, item in enumerate(all_items, 1):
        item["Item"] = i

    # Garante a ordem das colunas: Item, Quantidade, Informacao
    df_out = pd.DataFrame(all_items)
    df_out = df_out[["Item", "Quantidade", "Informacao"]]

    output_path = os.path.join(output_folder, excel_file)
    df_out.to_excel(output_path, index=False)
    print(f"[{excel_file}] Ficheiro criado em: {output_path}")

print("\nFim do processamento.")
