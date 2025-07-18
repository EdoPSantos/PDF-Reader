import os
import re
import json
import string
import unicodedata
import pandas as pd
from collections import defaultdict


import pandas as pd
import json
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

def is_quantity_number(text):
    try:
        t = str(text).replace(",", ".").strip()
        float(re.findall(r"[-+]?\d*[\.,]?\d+", t)[0])
        return True
    except:
        return False

def sort_items_by_key(items, *keys, default=0):
    return sorted(
        items,
        key=lambda item: tuple(item.get(k, default) for k in keys)
    )

def join_line_texts(words, filter_set=None):
    return " ".join(
        word["text"].strip()
        for word in words
        if word["text"].strip() and (filter_set is None or word["text"].strip() in filter_set)
    )

def filter_new_words_ordered(words_sorted, *texts_to_exclude):
    exclude_words = []
    for txt in texts_to_exclude:
        exclude_words += txt.split()
    
    exclude_counter = {}
    for w in exclude_words:
        exclude_counter[w] = exclude_counter.get(w, 0) + 1

    result = []
    for w in words_sorted:
        txt = w["text"].strip()
        if not txt:
            continue
        count = exclude_counter.get(txt, 0)
        if count > 0:
            exclude_counter[txt] -= 1
        else:
            result.append(w)
    return result

#------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------JSON----------------------------------------------------------------

def extract_items_json(pages, y_margin_possible_values=None):
    y_margin_line_values = 5
    header_names = []
    header_cord = []
    all_words = []
    json_items = []
    fallback_items = []
    normalized_keywords = [normalize_info(k) for k in QUANTITY_KEYWORDS]

    # -----------------------------------------------------------------------------------------------
    # ------------------------------------- Guardar Info Headers ------------------------------------

    # Junta a deteção do cabeçalho e a cópia das palavras com número da página
    for page_num, page in enumerate(pages, 1):
        for word in page["words"]:
            word_copy = dict(word)
            word_copy["page"] = page_num
            all_words.append(word_copy)

            if normalize_info(word["text"]) in normalized_keywords:
                header_names.append(word["text"])
                header_cord.append({"x": word["x"],
                                    "y": word["y"],
                                    "width": word.get("width", 0),
                                    "page": page_num,
                                    "text": word["text"]
                                    })

    # -----------------------------------------------------------------------------------------------
    # ---------------------------------------- Organizar Valores ------------------------------------

    # Agrupar palavras por página e coordenada y (linha)
    lines_by_page = defaultdict(lambda: defaultdict(list))
    for word in all_words:
        lines_by_page[word["page"]][int(round(word["y"]))].append(word)

    header_cord_by_page = defaultdict(list)
    for qc in header_cord:
        header_cord_by_page[qc["page"]].append((qc["x"], qc["y"]))

    for page in sorted(lines_by_page):
        qty_lines = []
        last_possible_values = ""
        lines = lines_by_page[page]
        found_valid_quantity = False
        y_sorted = sorted(lines.keys())

        # -----------------------------------------------------------------------------------------------
        # -------------------------------------- Guardar Coordenadas ------------------------------------

        headers_coords = [qc for qc in header_cord if qc["page"] == page]

        # -----------------------------------------------------------------------------------------------
        # ------------------------------------- Procura y dos Valores -----------------------------------

        # Encontrar todas as linhas que têm quantidades (nova abordagem)
        for y in y_sorted:
            words_in_line = lines[y]
            matching_words = []

            header_ranges = []
            for c in headers_coords:
                x_start = int(round(c["x"]))
                x_end = int(round(c["x"] + c.get("width", 0)))
                y_value = c["y"]

                header_ranges.append({
                    "x_range": set(range(x_start, x_end + 1)),
                    "y": y_value,
                    "text": c.get("text", "")
                })

            for word in words_in_line:
                word_x_range = set(range(
                    int(round(word["x"])),
                    int(round(word["x"] + word.get("width", 0))) + 1
                ))
                word_y = word["y"]

                for header in header_ranges:
                    if word_x_range & header["x_range"] and word_y > header["y"]:
                        if is_quantity_number(word["text"]):
                            matching_words.append(word)
                            break

            if matching_words:
                qty_lines.append(y)

        # Se não encontrou nenhuma linha de quantidade, faz fallback para todas as palavras do cabeçalho
        if not found_valid_quantity and headers_coords:
            # Tenta encontrar o y mais próximo da linha do cabeçalho
            header_y_raw = headers_coords[0]["y"]
            header_y = min(lines, key=lambda ly: abs(ly - header_y_raw))

            header_line_words = sort_items_by_key(lines[header_y], "x")
            header_cols = []
            for header_word in header_line_words:
                col_x = int(round(header_word["x"]))
                col_width = int(round(header_word.get("width", 0)))
                col_x_range = set(range(col_x, col_x + col_width + 1))
                header_cols.append({
                    "text": header_word["text"],
                    "x_range": col_x_range,
                    "y": header_word["y"]
                })

            for y in y_sorted:
                if y <= header_y:
                    continue
                found_cols = []
                for col in header_cols:
                    for word in lines[y]:
                        word_x = int(round(word["x"]))
                        word_width = int(round(word.get("width", 0)))
                        word_x_range = set(range(word_x, word_x + word_width + 1))
                        if word_x_range & col["x_range"]:
                            found_cols.append(col["text"])
                            break

                # Permite margem mínima: se pelo menos 2 colunas diferentes foram identificadas
                if len(set(found_cols)) >= 2:
                    # Controle de distância entre linhas
                    if 'last_y' not in locals():
                        # Primeiro valor após cabeçalho
                        first_y = y
                        last_y = y
                        line_words_sorted = sort_items_by_key(lines[y], "x")
                        no_line_values = join_line_texts(line_words_sorted)
                        no_target_quantity = "Não Identificado"
                        fallback_items.append({
                            "Quantity":  no_target_quantity,
                            "All Values": no_line_values,
                            "page": page,
                            "y": y,
                        })
                    else:
                        # Para os próximos valores, aplica o filtro de distância
                        dist_cabecalho_primeiro = abs(first_y - header_y)
                        dist_entre_valores = abs(y - last_y)
                        if dist_entre_valores < 2 * dist_cabecalho_primeiro:
                            last_y = y
                            line_words_sorted = sort_items_by_key(lines[y], "x")
                            line_values = join_line_texts(line_words_sorted)
                            no_target_quantity = "Não Identificado"
                            fallback_items.append({
                                "Quantity":  no_target_quantity,
                                "All Values": no_line_values,
                                "page": page,
                                "y": y,
                            })
        
        # -----------------------------------------------------------------------------------------------
        # ----------------------------------------- Define Margens --------------------------------------

        # Calcular Media da Distancia entre Valores
        if y_margin_possible_values is None:
            if len(qty_lines) > 1:
                diffs = [qty_lines[i+1] - qty_lines[i] for i in range(len(qty_lines)-1)]
                y_margin_possible_values_page = int(sum(diffs) / len(diffs))
            else:
                y_margin_possible_values_page = 30  # valor padrão
        else:
            y_margin_possible_values_page = y_margin_possible_values
        
        # -----------------------------------------------------------------------------------------------

        # Para cada quantidade, apanha a linha + as N linhas abaixo, sem apanhar a próxima quantidade
        for idx, y in enumerate(qty_lines):
            line_words = []
            possible_value_words = []
            block_lines = []
            header_ranges = []
            matching_words = []

            # -----------------------------------------------------------------------------------------------
            # ---------------------------------- Range para Recolha de Info ---------------------------------

            # Descobre até onde pode apanhar (sem incluir próxima quantidade)
            if idx + 1 < len(qty_lines):
                next_qty_y = qty_lines[idx + 1]
                valid_correct_ys = [yy for yy in y_sorted if y < yy < next_qty_y and (yy - y) <= y_margin_possible_values_page]
                block_ys = [yy for yy in y_sorted if y <= yy < next_qty_y]
            else:
                valid_correct_ys = [yy for yy in y_sorted if yy > y and (yy - y) <= y_margin_possible_values_page]
                block_ys = [yy for yy in y_sorted if yy >= y]

            # -----------------------------------------------------------------------------------------------
            # -------------------------------------- Linhas dos Valores -------------------------------------

            # Adiciona margem em y para capturar mais informação em line_values
            line_ys = [yy for yy in y_sorted if abs(yy - y) <= y_margin_line_values]
            
            for yy in line_ys:
                line_words.extend(lines[yy])
            line_words_sorted = sort_items_by_key(line_words, "x")
            line_values = join_line_texts(line_words_sorted)


            # -----------------------------------------------------------------------------------------------
            # ---------------------------------- Identificação de Limites -----------------------------------

            for header in headers_coords:
                header_y = int(round(header["y"]))
                if header_y in lines:
                    header_line_words = sort_items_by_key(lines[header_y], "x")
                    for i, word in enumerate(header_line_words):
                        left = header_line_words[i-1] if i > 0 else None
                        right = header_line_words[i+1] if i < len(header_line_words)-1 else None
                        left_xw = left["x"] + left.get("width", 0) if left else None
                        right_xw = right["x"] + right.get("width", 0) if right else None

                        if left_xw is not None and right_xw is not None:
                            delta_x = right_xw - left_xw

            
            # -----------------------------------------------------------------------------------------------
            # --------------------------------------- Possíveis Valores -------------------------------------

            # Junta todas as linhas dentro da margem vertical
            for yy in valid_correct_ys:
                possible_value_words.extend(lines[yy])
            possible_value_words_sorted = sort_items_by_key(possible_value_words, "x")
            all_text = join_line_texts(possible_value_words_sorted)

            # Filtra palavras já vistas
            filtered_words = filter_new_words_ordered(possible_value_words_sorted, line_values, last_possible_values)
            possible_values = join_line_texts(filtered_words)

            # Só atualiza o histórico se possível_values tiver conteúdo
            if possible_values.strip():
                last_possible_values = all_text

            # -----------------------------------------------------------------------------------------------
            # ------------------------------------- Quantidades-Valores -------------------------------------

            for c in headers_coords:
                x_start = int(round(c["x"]))
                x_end = int(round(c["x"] + c.get("width", 0)))
                y_value = c["y"]

                header_ranges.append({
                    "x_range": set(range(x_start, x_end + 1)),
                    "y": y_value,
                    "text": c.get("text", "")
                })

            for word in lines[y]:
                word_x_range = set(range(
                    int(round(word["x"])),
                    int(round(word["x"] + word.get("width", 0))) + 1
                ))
                word_y = word["y"]

                for header in header_ranges:
                    # Verifica se há intersecção no X E se está abaixo no Y
                    if word_x_range & header["x_range"] and word_y > header["y"]:
                        matching_words.append(word)
                        break  # Já pertence a um header, não precisa testar os outros

            matching_words = sort_items_by_key(matching_words, "y")
            target_quantities = [
                normalize_quantity(word["text"])
                for word in matching_words
                if word.get("width", 0) < delta_x
            ]
            target_quantity = " ".join(target_quantities)

            # -----------------------------------------------------------------------------------------------
            # ------------------------------------ Todos os Valores Sumados ---------------------------------
            
            all_words_combined = line_words_sorted + filtered_words
            all_words_sorted = sort_items_by_key(all_words_combined, "x")
            all_values = join_line_texts(all_words_sorted)

            # -----------------------------------------------------------------------------------------------
            # ---------------------------------- Todos os Valores 100% do PDF -------------------------------

            for yy in block_ys:
                block_lines.extend(lines[yy])
            block_lines = sort_items_by_key(block_lines, "x")
            full_content = join_line_texts(block_lines)

            # -----------------------------------------------------------------------------------------------
            # ----------------------------------------- Guarda Valores ---------------------------------------

            # Guarda o item
            if target_quantity.strip():
                found_valid_quantity = True
                json_items.append({
                    "Quantity": target_quantity,
                    "All Values": all_values,
                    "page": page,
                    "y": y,
                })

    if json_items:
        return json_items
    else:
        return fallback_items

#----------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------Main-----------------------------------------------------------------
def main():
    #------------------------------------------------------------------------
    # ------------------------------- 3. JSON -------------------------------

    jsons_folder = "../jsons"
    result_folder = os.path.join("..", "result")
    if not os.path.exists(jsons_folder):
        print("Pasta JSONs não encontrada.")
        return
    json_files = [f for f in os.listdir(jsons_folder) if f.endswith('.json')]
    if not json_files:
        print("Nenhum ficheiro JSON encontrado na pasta.")
        return

    # Create result directory once
    result_folder = os.path.join("..", "result")
    os.makedirs(result_folder, exist_ok=True)

    #------------------------------------------------------------------------
    # ------------------------ Ciclo de Processamento -----------------------

    for json_file in json_files:
        try:
            json_path = os.path.join(jsons_folder, json_file)
            json_items = []
            if os.path.exists(json_path):
                with open(json_path, encoding="utf-8") as f:
                    pages = json.load(f)
                json_items = extract_items_json(pages)
            else:
                print(f"JSON não encontrado: {json_path}")
                continue

            base_name = os.path.splitext(os.path.basename(json_path))[0]
            excel_out_path = os.path.join(result_folder, f"{base_name}.xlsx")

            #-----------------------------------------------------------------------
            #------------------------------ Items Value ----------------------------
           
            total_items = len(json_items)
            items_by_page = defaultdict(list)
            for idx, item in enumerate(json_items, 1):
                page = item.get('page', 'N/A')
                try:
                    page = int(page)
                except Exception:
                    pass
                items_by_page[page].append(item)

            #-----------------------------------------------------------------------
            #------------------------------- Statistics ----------------------------
            '''
            resumo_rows = []
            resumo_rows.append(["Página", "Produtos"])
            for page in sorted(items_by_page):
                resumo_rows.append([f"Pag: {page}", len(items_by_page[page])])
            resumo_rows.append(["Total Produtos", total_items])
            '''
            #-----------------------------------------------------------------------
            #------------------------------- Main Info -----------------------------

            detalhes_rows = []
            detalhes_rows.append(["Codigo", "Quantidade", "Observacoes"])
            for page in sorted(items_by_page):
                #detalhes_rows.append([f"Pag: {page}"])
                for item in items_by_page[page]:
                    detalhes_rows.append([
                        "",
                        item.get("Quantity", ""),
                        item.get("All Values", ""),
                    ])
                #detalhes_rows.append([])

            #-----------------------------------------------------------------------
            #------------------------- Statistics + Main Info ----------------------
            
            all_rows = detalhes_rows

            #-----------------------------------------------------------------------
            #------------------------------ Write Excel ----------------------------
            
            with pd.ExcelWriter(excel_out_path, engine='xlsxwriter') as writer:
                df = pd.DataFrame(all_rows)
                df.to_excel(writer, sheet_name='Itens', index=False, header=False)

        except Exception as e:
            print(f"Ocorreu um erro ao processar {json_file}: {e}")

#--------------------------------------------------------------------------------------------
#---------------------------------------Show-Main--------------------------------------------
if __name__ == "__main__":
    main()