import os
import pandas as pd
import re
import unicodedata
import json
import string
from collections import defaultdict

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

#------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------Global-Helpers--------------------------------------------------------------

QUANTITY_KEYWORDS = ["quantidade", "qtd", "qt", "quantity", "q"]

def normalize_info(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    # Remove acentuação
    text = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    # Remove pontuação (.,:;- etc.)
    text = ''.join(c for c in text if c not in string.punctuation)
    return text

def normalize_quantity(q):
    try:
        return str(int(float(str(q).replace(",", ".").strip())))
    except:
        return str(q).strip()

def is_garbage(text):
    norm = normalize_info(text)
    return norm == "" or norm == "nan"

def is_quantity_number(text):
    try:
        t = str(text).replace(",", ".").strip()
        float(re.findall(r"[-+]?\d*[\.,]?\d+", t)[0])
        return True
    except:
        return False
    
def get_file_title_from_path(path):
    """Retorna o título do tipo de ficheiro a partir do caminho."""
    if not path:
        return ""
    ext = os.path.splitext(str(path))[-1].lower()
    if ext == ".json":
        return "JSON"
    elif ext in [".xlsx", ".xls"]:
        return "Excel"
    elif ext == ".pdf":
        return "PDF"
    else:
        return os.path.basename(str(path))
    
def item_per_page(title, items, order_key=None):
    """Imprime os itens agrupados por página, usando print_itens para cada página."""
    pages = defaultdict(list)
    for item in items:
        page = item.get("page", 1)
        pages[page].append(item)
    for page in sorted(pages):
        page_info = f"{title} - Página {page} de {len(pages)}"
        print_items(title, pages[page], page_info, order_key=order_key)

def print_items(title, items, pages, order_key=None):
    print(f"\n--- Resultados do ficheiro: {pages} tem {len(items)} itens ---")
    if not items:
        print("Nenhum item válido encontrado.")
        return
    if order_key and all(order_key in item for item in items):
        items = sorted(items, key=lambda x: x[order_key])
    for i, item in enumerate(items, 1):
        print(f"\n {title} - Item {i}:")
        print(f"  Quantidade: {item['Quantidade']}")
        print(f"  Informação: {item['Informacao']}")

def sort_items_by_key(items, *keys, default=0):
    return sorted(
        items,
        key=lambda item: tuple(item.get(k, default) for k in keys)
    )

def join_line_texts(words):
    return " ".join(word["text"].strip() for word in words if word["text"].strip())
    
#------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------Excel------------------------------------------------------------------
# ----- EXCEL: mantém a lógica original -----
def find_headers(df, keywords):
    n_rows, n_cols = df.shape
    header_row = None
    cols_found = []
    for row in range(n_rows):
        for col in range(n_cols):
            cell = normalize_info(df.iat[row, col])
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
                    if not is_garbage(val) and is_quantity_number(val):
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

#------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------Filters-----------------------------------------------------------------

def distance_to_integer(q):
    try:
        return abs(float(q) - round(float(q)))
    except:
        return 9999
    
def filter_duplicates_by_info(items):
    filtered = {}
    for item in items:
        # Usa o campo novo
        info_key = normalize_info(item.get('Possiveis valores', ''))
        quant = item['Quantidade']
        # Se já existe, compara a proximidade ao inteiro
        if info_key in filtered:
            q_current = filtered[info_key]['Quantidade']
            if distance_to_integer(quant) < distance_to_integer(q_current):
                filtered[info_key] = item
        else:
            filtered[info_key] = item
    return list(filtered.values())

#------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------???-------------------------------------------------------------------
# ----- JSON/PDF: blocos entre quantidades -----
def extract_blocks_json_or_pdf(lines):
    y_sorted = sorted(lines.keys())
    qty_idxs = []
    qty_vals = []
    qty_line_indices = []
    for i, y in enumerate(y_sorted):
        words_sorted = sorted(lines[y], key=lambda w: w["x"])
        for w in words_sorted:
            if is_quantity_number(w["text"]):
                qty_idxs.append(i)
                qty_line_indices.append(y)
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
        # "Valores corretos" = apenas a linha onde está a quantidade
        q_line_words = sorted(lines[y_sorted[start_idx]], key=lambda w: w["x"])
        valores_corretos = " ".join(w["text"] for w in q_line_words)
        items.append({
            "Quantidade": qty_vals[idx],
            "Valores corretos": valores_corretos,
            "Possiveis valores": " ".join(block_lines)
        })
    return items

#------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------------------PDF-------------------------------------------------------------------
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

#------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------JSON----------------------------------------------------------------

def extract_items_json(pages, x_margin=20, y_margin_possible_values=30):
    y_margin_line_values = 5
    header_x = None
    header_y = None
    header_x_1 = None
    header_y_1 = None
    quant_names = []
    quant_cord = []
    all_words = []
    json_items = []
    normalized_keywords = [normalize_info(k) for k in QUANTITY_KEYWORDS]

    # Junta a deteção do cabeçalho e a cópia das palavras com número da página
    for page_num, page in enumerate(pages, 1):
        for word in page["words"]:
            word_copy = dict(word)
            word_copy["page"] = page_num
            all_words.append(word_copy)

            if normalize_info(word["text"]) in normalized_keywords:
                quant_names.append(word["text"])
                quant_cord.append({"x": word["x"], "y": word["y"], "page": page_num})
                if header_x_1 is None:
                    header_x_1 = word["x"]
                    header_y_1 = word["y"]

    # Agrupar palavras por página e coordenada y (linha)
    lines_by_page = defaultdict(lambda: defaultdict(list))
    for word in all_words:
        lines_by_page[word["page"]][int(round(word["y"]))].append(word)

    quant_cord_by_page = defaultdict(list)
    for qc in quant_cord:
        quant_cord_by_page[qc["page"]].append((qc["x"], qc["y"]))

    for page in sorted(lines_by_page):
        qty_lines = []
        lines = lines_by_page[page]
        y_sorted = sorted(lines.keys())

        page_quant_coords = quant_cord_by_page.get(page, [])

        page_coords = [qc for qc in quant_cord if qc["page"] == page]
        if page_coords:
            header_x = page_coords[0]["x"]
            header_y = page_coords[0]["y"]
        else:
            header_x = None
            header_y = None

        # Encontrar todas as linhas que têm quantidades
        for y in y_sorted:
            words_in_line = lines[y]
            # Para cada coordenada de quantidade nesta página
            for qx, qy in page_quant_coords:
                qty_num = next(
                    (word for word in lines[y]
                    if is_quantity_number(word["text"])
                    and abs(word["x"] - header_x) <= x_margin
                    and word["y"] > header_y),
                    None
                )
                if qty_num:
                    qty_lines.append(y)
                    break  # Só precisa de um match por linha
        if not qty_lines:
            continue

        # Para cada quantidade, apanha a linha + as N linhas abaixo, sem apanhar a próxima quantidade
        for idx, y in enumerate(qty_lines):
            line_words = []
            possible_value_words = []
            block_lines = []
            # Descobre até onde pode apanhar (sem incluir próxima quantidade)
            if idx + 1 < len(qty_lines):
                next_qty_y = qty_lines[idx + 1]
                valid_correct_ys = [yy for yy in y_sorted if yy < next_qty_y and abs(yy - y) <= y_margin_possible_values]
                block_ys = [yy for yy in y_sorted if y <= yy < next_qty_y]
            else:
                valid_correct_ys = [yy for yy in y_sorted if abs(yy - y) <= y_margin_possible_values]
                block_ys = [yy for yy in y_sorted if y <= yy]
            
            # Adiciona margem em y para capturar mais informação em line_values
            line_ys = [yy for yy in y_sorted if abs(yy - y) <= y_margin_line_values]
            
            for yy in line_ys:
                line_words.extend(lines[yy])
            line_words_sorted = sort_items_by_key(line_words, "x")
            line_values = join_line_texts(line_words_sorted)
            
            # ------------------------------------------- ALTERAR -------------------------------------------
            # Atualmente procura linhas expecificas, mas o valor deve ser procurado atravez da subtração dos valores do "line_values"
            
            for yy in valid_correct_ys:
                possible_value_words.extend(lines[yy])
            possible_value_words_sorted = sort_items_by_key(possible_value_words, "x")
            possible_values = join_line_texts(possible_value_words_sorted)
            # -----------------------------------------------------------------------------------------------

            qty_num = next(
                (word for word in lines[y]
                if is_quantity_number(word["text"])
                and abs(word["x"] - header_x) <= x_margin
                and word["y"] > header_y),
                None
            )
            target_quantity = normalize_quantity(qty_num["text"]) if qty_num else ""

            for yy in block_ys:
                block_lines.extend(lines[yy])
            block_lines = sort_items_by_key(block_lines, "x")
            full_content = join_line_texts(block_lines)
            
            # Guarda o item
            json_items.append({
                "Keywords quantidade": quant_names,
                "Quantidades coordenadas": quant_cord,
                "Quantidade": target_quantity,
                "Valores da linha": line_values,
                "Possiveis valores": possible_values,
                "Conteúdo total": full_content,
                "page": page,
                "y": y
            })

    return json_items

#----------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------Main-----------------------------------------------------------------
def main():
    
    #-------------------------------------------------------------------
    # ---------------------------- 1. Excel ----------------------------

    excel_path = "../excels_visual/opcao3.xlsx"
    all_excel_items = []
    excel_quantidades = set()
    if os.path.exists(excel_path):
        sheets = pd.read_excel(excel_path, header=None, sheet_name=None)
        for sheet_name, df in sheets.items():
            excel_items = extract_items_quantity_only(df)
            excel_items = filter_duplicates_by_info(excel_items)
            all_excel_items.extend(excel_items)
            file_name_excel = get_file_title_from_path(excel_path)
    else:
        print("Excel não encontrado.")
        return

    #-----------------------------------------------------------------
    # ---------------------------- 2. PDF ----------------------------

    pdf_path = "../pdfs_out/opcao3.pdf"
    pdf_items = []
    if os.path.exists(pdf_path):
        pdf_items = extract_items_pdf(pdf_path)
        pdf_items = filter_duplicates_by_info(pdf_items)
        file_name_pdf = get_file_title_from_path(pdf_path)
    else:
        print("PDF não encontrado.")

    #------------------------------------------------------------------
    # ---------------------------- 3. JSON ----------------------------

    jsons_folder = "../jsons"
    if not os.path.exists(jsons_folder):
        print("Pasta JSONs não encontrada.")
        return
    json_files = [f for f in os.listdir(jsons_folder) if f.endswith('.json')]
    if not json_files:
        print("Nenhum ficheiro JSON encontrado na pasta.")
        return

    #-------------------------------------------------------------------
    # ------------------------ Imprime a Versão ------------------------
    # Calcule o contador antes do loop
    counter_path = os.path.join("..", "result", "contador.txt")
    if os.path.exists(counter_path):
        with open(counter_path, "r", encoding="utf-8") as f:
            try:
                counter = int(f.read().strip())
            except Exception:
                counter = 0
    else:
        counter = 0
    counter += 1

    erro_ocorrido = False

    #--------------------------------------------------------------------
    # ---------------------- Ciclo de Processamento ---------------------
    for json_file in json_files:
        try:
            json_path = os.path.join(jsons_folder, json_file)
            json_items = []
            if os.path.exists(json_path):
                with open(json_path, encoding="utf-8") as f:
                    pages = json.load(f)
                json_items = extract_items_json(pages)
                json_items = filter_duplicates_by_info(json_items)
                file_name_json = get_file_title_from_path(json_path)
            else:
                print(f"JSON não encontrado: {json_path}")
                continue

            if not os.path.exists(os.path.join("..", "result")):
                print("Criando pasta de resultados...")

            os.makedirs(os.path.join("..", "result"), exist_ok=True)

            base_name = os.path.splitext(os.path.basename(json_path))[0]
            excel_out_path = os.path.join("..", "result", f"{base_name}.xlsx")

            txt_file = os.path.join("..", "result", f"{base_name}.txt")

            export_rows = []
            for idx, item in enumerate(json_items, 1):
                export_rows.append({
                    "Item": idx,
                    "Página": item.get("page", "N/A"),
                    "Quantidade": item.get("Quantidade", ""),
                    "Valores corretos": item.get("Valores corretos", ""),
                    "Possiveis valores": item.get("Possiveis valores", ""),
                    "Conteúdo total": item.get("Conteúdo total", ""),
                })

    #----------------------------------------------------------------------------------------------------------------------------------
    #----------------------------------------------------------Teste-Zone--------------------------------------------------------------

            # ---------------- Imprime os dados num txt para melhor análise de dados ----------------
            # Contador automático de execuções
            # Agrupa itens por página para imprimir uma vez por página
            items_by_page = defaultdict(list)
            for idx, item in enumerate(json_items, 1):
                items_by_page[item.get('page', 'N/A')].append((idx, item))

            
            txt_out_path = os.path.join("..", "result", f"{base_name}.txt")
            with open(txt_out_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(f"Versão: {counter}\n\n")
                if json_items:
                    txt_file.write(f"Keywords de quantidade encontradas: {', '.join(json_items[0].get('Keywords quantidade', []))}\n\n")
                    txt_file.write(f"Quantidades coordenadas: {', '.join(str(coord) for coord in json_items[0].get('Quantidades coordenadas', []))}\n\n")

                #txt_file.write(f"Cuidado!!!\n")
                #txt_file.write(f"A Informação pode estar incompleta ou incorreta.\n")
                #txt_file.write(f"Precauções:.\n")
                #txt_file.write(f"O campo (Valores corretos) pode ter valores a mais.\n")
                #txt_file.write(f"O campo (Possíveis valores) pode estar incompleto ou misturar valores.\n")
                #txt_file.write(f"O campo (Conteúdo total) pode pode ter valores a mais e misturar valores.\n")
                for page in sorted(items_by_page):
                    txt_file.write(f"Página {page}:\n")
                    for idx, item in items_by_page[page]:
                        txt_file.write(f"  Item: {idx}\n")
                        txt_file.write(f"    Quantidade: {item.get('Quantidade', '')}\n")
                        txt_file.write(f"    Valores da linha: {item.get('Valores da linha', '')}\n")
                        txt_file.write(f"    Possiveis valores: {item.get('Possiveis valores', '')}\n")
                        txt_file.write(f"    Conteúdo total: {item.get('Conteúdo total', '')}\n")
                        txt_file.write("-" * 40 + "\n")

            df_export = pd.DataFrame(export_rows)
            df_export.to_excel(excel_out_path, index=False)
            print(f"\nResultados exportados para {excel_out_path}")

        except Exception as e:
            print(f"Ocorreu um erro ao processar {json_file}: {e}")
            erro_ocorrido = True        
    #-------------------------------------------------------------------------------------------
    # ---------------- Imprime os dados todos comparando com o json com o excel ----------------

    #print(f"\n--- Resultados do ficheiro {file_name_json}: ---")
    #print(f"Total de itens: {len(json_items)}")

    #pagina_atual = None
    #for item in json_items:
    #    page = item.get('page', 'N/A')
    #    if page != pagina_atual:
    #        print(f"\nPágina {page}:")
    #        pagina_atual = page
    #    quant = item.get('Quantidade', 'N/A')
    #    info = item.get('Informacao', 'N/A')
    #    print(f"  Quantidade: {quant}, Informação: {info}")

    #---------------------------------------------------------
    # ---------------- Confirma se houve erro ----------------

    # Só atualiza o contador se tudo der certo
    if not erro_ocorrido:
        with open(counter_path, "w", encoding="utf-8") as f:
            f.write(str(counter))
#----------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------Show-Main---------------------------------------------------------------
if __name__ == "__main__":
    main()