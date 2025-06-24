import os
import pandas as pd
import re
import unicodedata
import json

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

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

def normalize_quantity(q):
    try:
        return str(int(float(str(q).replace(",", ".").strip())))
    except:
        return str(q).strip()
    
def distance_to_integer(q):
    try:
        return abs(float(q) - round(float(q)))
    except:
        return 9999

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

# ----- EXCEL: mantém a lógica original -----
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
    already_found = set()
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
                            quantidade_int = val
                        items.append({
                            "Quantidade": quantidade_int,
                            "Informacao": info
                        })
                        break
    return items

def filter_duplicates_by_info(items):
    filtered = {}
    for item in items:
        info_key = normalize(item['Informacao'])
        quant = item['Quantidade']
        # Se já existe, compara a proximidade ao inteiro
        if info_key in filtered:
            q_current = filtered[info_key]['Quantidade']
            if distance_to_integer(quant) < distance_to_integer(q_current):
                filtered[info_key] = item
        else:
            filtered[info_key] = item
    return list(filtered.values())

def align_blocks_by_excel(excel_items, other_items):
    aligned = []
    # Constrói dict {quantidade normalizada: item} para pesquisa rápida no PDF/JSON
    lookup = {normalize_quantity(it["Quantidade"]): it for it in other_items}
    for item in excel_items:
        q_norm = normalize_quantity(item["Quantidade"])
        other_item = lookup.get(q_norm, None)
        aligned.append((q_norm, item["Informacao"], other_item["Informacao"] if other_item else None))
    return aligned

def print_aligned_items(title, excel_items, other_items):
    aligned = align_blocks_by_excel(excel_items, other_items)
    print(f"\n--- {title} ALINHADO COM EXCEL ---")
    for idx, (q, info_excel, info_other) in enumerate(aligned, 1):
        print(f"\nItem {idx}:")
        print(f"  Quantidade: {q}")
        print(f"  Info Excel: {info_excel}")
        if info_other:
            print(f"  Info {title}: {info_other}")
        else:
            print(f"  Info {title}: Não encontrado")

# ----- JSON/PDF: blocos entre quantidades -----
def extract_blocks_json_or_pdf(lines):
    # Ordena linhas por Y (de cima para baixo)
    y_sorted = sorted(lines.keys())
    # Lista das posições das linhas com quantidade (âncora)
    qty_idxs = []
    qty_vals = []
    for i, y in enumerate(y_sorted):
        words_sorted = sorted(lines[y], key=lambda w: w["x"])
        for w in words_sorted:
            if is_quantity_number(w["text"]):
                qty_idxs.append(i)
                try:
                    qty_val = int(float(w["text"].replace(",", ".")))
                except Exception:
                    qty_val = w["text"]
                qty_vals.append(qty_val)
                break
    if not qty_idxs:
        return []

    items = []
    for idx, start_idx in enumerate(qty_idxs):
        end_idx = qty_idxs[idx + 1] if idx + 1 < len(qty_idxs) else len(y_sorted)
        block_lines = []
        for i in range(start_idx, end_idx):
            y = y_sorted[i]
            words_sorted = sorted(lines[y], key=lambda w: w["x"])
            line_text = " ".join(w["text"] for w in words_sorted)
            block_lines.append(line_text)
        items.append({
            "Quantidade": qty_vals[idx],
            "Informacao": " ".join(block_lines)
        })
    return items

def extract_items_pdf(path):
    if not fitz:
        print("PyMuPDF não está instalado.")
        return []
    doc = fitz.open(path)
    all_items = []
    for page_num, page in enumerate(doc, 1):
        blocks = page.get_text("dict")["blocks"]
        words = []
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    angle = span.get("angle", 0)
                    if abs(angle) > 5:
                        continue
                    words.append({
                        "text": span["text"],
                        "x": span["bbox"][0],
                        "y": span["bbox"][1],
                    })
        lines = {}
        for word in words:
            y = int(round(word["y"] / 3))
            lines.setdefault(y, []).append(word)
        items = extract_blocks_json_or_pdf(lines)
        all_items.extend(items)
    return all_items

def extract_items_json(pages, x_margin=5):
    all_items = []
    for page in pages:
        lines = {}
        for word in page["words"]:
            y = int(round(word["y"] / 3))
            lines.setdefault(y, []).append(word)
        # 1. Encontra o cabeçalho e X da coluna quantidade
        header_y, header_x = None, None
        for y, words in lines.items():
            for w in words:
                if normalize(w["text"]) in QUANTITY_KEYWORDS:
                    header_y = y
                    header_x = w["x"]
                    break
            if header_x is not None:
                break
        if header_x is None:
            continue  # Se não encontrou cabeçalho, passa à próxima página

        # 2. Para as linhas seguintes, procura quantidade só na coluna certa
        for y in sorted(lines):
            if y <= header_y: continue
            words_sorted = sorted(lines[y], key=lambda w: w["x"])
            quantidade = None
            for w in words_sorted:
                # Aceita apenas valores com X até ±5 do header_x
                if abs(w["x"] - header_x) <= x_margin and is_quantity_number(w["text"]):
                    quantidade = normalize_quantity(w["text"])
                    break
            if quantidade is not None:
                line_text = " ".join(w["text"] for w in words_sorted)
                all_items.append({
                    "Quantidade": quantidade,
                    "Informacao": line_text
                })
    return all_items

def print_itens(titulo, items):
    print(f"\n--- Resultados do ficheiro: {titulo} ---")
    if not items:
        print("Nenhum item válido encontrado.")
        return
    for i, item in enumerate(items, 1):
        print(f"\nItem {i}:")
        print(f"  Quantidade: {item['Quantidade']}")
        print(f"  Informação: {item['Informacao']}")

def main():
    # 1. Excel
    excel_path = "../excels_visual/opcao3.xlsx"
    all_excel_items = []
    excel_quantidades = set()
    if os.path.exists(excel_path):
        sheets = pd.read_excel(excel_path, header=None, sheet_name=None)
        for sheet_name, df in sheets.items():
            items = extract_items_quantity_only(df)
            items = filter_duplicates_by_info(items)
            for item in items:
                excel_quantidades.add(normalize_quantity(item["Quantidade"]))
            all_excel_items.extend(items)
    else:
        print("Excel não encontrado.")
        return

    # 2. PDF
    pdf_path = "../pdfs_out/opcao3.pdf"
    pdf_items = []
    if os.path.exists(pdf_path):
        pdf_items = extract_items_pdf(pdf_path)
        pdf_items = filter_duplicates_by_info(pdf_items)
    else:
        print("PDF não encontrado.")

    # 3. JSON
    json_path = "../jsons/opcao3.json"
    json_items = []
    if os.path.exists(json_path):
        with open(json_path, encoding="utf-8") as f:
            pages = json.load(f)
        json_items = extract_items_json(pages, x_margin=5)
        json_items = filter_duplicates_by_info(json_items)
    else:
        print("JSON não encontrado.")

    # --- COMPARAÇÃO FINAL ---
    pdf_lookup = {normalize_quantity(item["Quantidade"]): item["Informacao"] for item in pdf_items}
    json_lookup = {normalize_quantity(item["Quantidade"]): item["Informacao"] for item in json_items}

    print("\n--- Comparação por quantidade (ordem do Excel) ---")
    for idx, item in enumerate(all_excel_items, 1):
        q_norm = normalize_quantity(item["Quantidade"])
        info_excel = item["Informacao"]
        info_pdf = pdf_lookup.get(q_norm, "Não encontrado")
        info_json = json_lookup.get(q_norm, "Não encontrado")
        print(f"\nItem {idx}:")
        print(f"  Quantidade: {q_norm}")
        print(f"  Info Excel: {info_excel}")
        print(f"  Info PDF: {info_pdf}")
        print(f"  Info JSON: {info_json}")

if __name__ == "__main__":
    main()