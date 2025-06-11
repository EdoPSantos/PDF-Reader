import os
import pdfplumber
import pandas as pd

input_folder = "pdfs"
output_folder = "excels_from_pdf"
os.makedirs(output_folder, exist_ok=True)

# Define as divisões horizontais da "grelha invisível" (em pontos PDF)
# Exemplo: 4 colunas de 150px numa largura total de 600px
colunas_x = [0, 20, 40, 60, 80, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600]

def extrair_com_grelha(page, colunas_x):
    words = page.extract_words()
    if not words:
        return []

    linhas = {}
    for word in words:
        y = round(word["top"])  # Agrupamento vertical
        x = word["x0"]

        # Identifica a que coluna da grelha pertence
        coluna_idx = None
        for i in range(len(colunas_x) - 1):
            if colunas_x[i] <= x < colunas_x[i + 1]:
                coluna_idx = i
                break

        if coluna_idx is not None:
            if y not in linhas:
                linhas[y] = [""] * (len(colunas_x) - 1)
            linhas[y][coluna_idx] += word["text"] + " "

    # Ordena por posição vertical (de cima para baixo)
    linhas_ordenadas = [linhas[y] for y in sorted(linhas)]
    return linhas_ordenadas

def pdf_para_excel(pdf_path, excel_path):
    all_tables = []
    print(f"A processar: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            grelha = extrair_com_grelha(page, colunas_x)
            if grelha:
                df = pd.DataFrame(grelha)
                df['Página'] = page_num
                all_tables.append(df)

    if all_tables:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for idx, table in enumerate(all_tables):
                table.to_excel(writer, sheet_name=f"Pag_{idx+1}", index=False, header=False)
        print(f"Criado: {excel_path}")
    else:
        print(f"Nenhum conteúdo relevante encontrado: {pdf_path}")

# === Loop por todos os PDFs ===
for filename in os.listdir(input_folder):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(input_folder, filename)
        excel_name = os.path.splitext(filename)[0] + ".xlsx"
        excel_path = os.path.join(output_folder, excel_name)

        if os.path.exists(excel_path):
            print(f"Ignorado (já existe): {excel_name}")
            continue

        pdf_para_excel(pdf_path, excel_path)
