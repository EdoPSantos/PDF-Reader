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
    header_x_1 = None
    header_y_1 = None
    header_width = None
    quant_names = []
    quant_cord = []
    all_words = []
    json_items = []
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
                quant_names.append(word["text"])
                quant_cord.append({"x": word["x"],
                                    "y": word["y"],
                                    "width": word.get("width", 0),
                                    "page": page_num})    
                if header_x_1 is None:
                    header_x_1 = word["x"]
                    header_y_1 = word["y"]
                    header_width = word.get("width", 0)

    # -----------------------------------------------------------------------------------------------
    # ---------------------------------------- Organizar Valores ------------------------------------

    # Agrupar palavras por página e coordenada y (linha)
    lines_by_page = defaultdict(lambda: defaultdict(list))
    for word in all_words:
        lines_by_page[word["page"]][int(round(word["y"]))].append(word)

    quant_cord_by_page = defaultdict(list)
    for qc in quant_cord:
        quant_cord_by_page[qc["page"]].append((qc["x"], qc["y"]))

    for page in sorted(lines_by_page):
        qty_lines = []
        last_possible_values = ""
        lines = lines_by_page[page]
        y_sorted = sorted(lines.keys())

        # -----------------------------------------------------------------------------------------------
        # -------------------------------------- Guardar Coordenadas ------------------------------------

        page_coords = [qc for qc in quant_cord if qc["page"] == page]
        
        # -----------------------------------------------------------------------------------------------
        # ------------------------------------- Procura y dos Valores -----------------------------------

        # Encontrar todas as linhas que têm quantidades
        for y in y_sorted:
            words_in_line = lines[y]
            qty_num = next(
                (word for word in words_in_line
                if is_quantity_number(word["text"])
                and any(
                    qc["x"] <= word["x"] <= qc["x"] + qc.get("width", 0)
                    for qc in page_coords
                )
                and word["y"] > min(qc["y"] for qc in page_coords)  # usa o menor y dos headers
                ),
                None
            )
            if qty_num:
                qty_lines.append(y)
        if not qty_lines:
            continue
        
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
            # ------------------------------------------ Quantidades ----------------------------------------

            # ------------------- Falta procurar a quantidade pela sua largura x + "width"
            qty_num = next(
                (word for word in lines[y]
                if is_quantity_number(word["text"])
                and any(
                    qc["x"] <= word["x"] <= qc["x"] + qc.get("width", 0)
                    for qc in page_coords
                )
                and word["y"] > min(qc["y"] for qc in page_coords)
                ),
                None
            )
            target_quantity = normalize_quantity(qty_num["text"]) if qty_num else ""

            # -----------------------------------------------------------------------------------------------
            # ---------------------------------------- Todos os Valores -------------------------------------
            
            for yy in block_ys:
                block_lines.extend(lines[yy])
            block_lines = sort_items_by_key(block_lines, "x")
            full_content = join_line_texts(block_lines)

            # -----------------------------------------------------------------------------------------------
            # ----------------------------------------- Guarda Valores ---------------------------------------

            # Guarda o item
            json_items.append({
                "Keywords quantidade": quant_names,
                "Quantidades coordenadas": quant_cord,
                "Quantidade": target_quantity,
                "Valores da linha": line_values,
                "Possiveis valores": possible_values,
                "Conteúdo total": full_content,
                "page": page,
                "y": y,
            })
            
            #print(f"Item adicionado: {json_items[-1]}")

    return json_items

#----------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------------------------------Main-----------------------------------------------------------------
def main():
    
    #------------------------------------------------------------------------
    # ------------------------------- 3. JSON -------------------------------

    jsons_folder = "../jsons"
    if not os.path.exists(jsons_folder):
        print("Pasta JSONs não encontrada.")
        return
    json_files = [f for f in os.listdir(jsons_folder) if f.endswith('.json')]
    if not json_files:
        print("Nenhum ficheiro JSON encontrado na pasta.")
        return

    #-------------------------------------------------------------------------
    # --------------------------- Imprime a Versão ---------------------------
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

    #________________________________________________________________________
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

            #--------------------------------------------------------------------------------------------
            #--------------------------------------- Teste-Zone -----------------------------------------

            # ---------------- Imprime os dados num txt para melhor análise de dados ----------------
            # Contador automático de execuções
            # Agrupa itens por página para imprimir uma vez por página
            items_by_page = defaultdict(list)
            for idx, item in enumerate(json_items, 1):
                page = item.get('page', 'N/A')
                try:
                    page = int(page)
                except Exception:
                    pass
                items_by_page[page].append((idx, item))

            
            txt_out_path = os.path.join("..", "result", f"{base_name}.txt")
            with open(txt_out_path, "w", encoding="utf-8") as txt_file:
                txt_file.write(f"Versão: {counter}\n\n")
                #txt_file.write(f"Cuidado!!!\n")
                #txt_file.write(f"A Informação pode estar incompleta ou incorreta.\n")
                #txt_file.write(f"Precauções:.\n")
                #txt_file.write(f"O campo (Valores corretos) pode ter valores a mais.\n")
                #txt_file.write(f"O campo (Possíveis valores) pode estar incompleto ou misturar valores.\n")
                #txt_file.write(f"O campo (Conteúdo total) pode pode ter valores a mais e misturar valores.\n")
                print("Páginas agrupadas:", list(items_by_page.keys()))
                for page in sorted(items_by_page, key=lambda x: (isinstance(x, int), x)):
                    txt_file.write(f"Página {page}:\n")
                    for idx, item in items_by_page[page]:
                        txt_file.write(f"  Item: {idx}\n")
                        txt_file.write(f"    Quantidade: {item.get('Quantidade', '')}\n")
                        txt_file.write(f"    Valores da linha: {item.get('Valores da linha', '')}\n")
                        txt_file.write(f"    Possiveis valores: {item.get('Possiveis valores', '')}\n")
                        txt_file.write(f"    Conteúdo total (Dividido por itens): {item.get('Conteúdo total', '')}\n")
                        txt_file.write("-" * 40 + "\n")

            df_export = pd.DataFrame(export_rows)
            df_export.to_excel(excel_out_path, index=False)
            print(f"\nResultados exportados para {excel_out_path}")

        except Exception as e:
            print(f"Ocorreu um erro ao processar {json_file}: {e}")
            erro_ocorrido = True        
   
    #-------------------------------------------------------------
    #___________________ Confirma se houve erro __________________

    # Só atualiza o contador se tudo der certo
    if not erro_ocorrido:
        with open(counter_path, "w", encoding="utf-8") as f:
            f.write(str(counter))

#--------------------------------------------------------------------------------------------
#---------------------------------------Show-Main--------------------------------------------
if __name__ == "__main__":
    main()