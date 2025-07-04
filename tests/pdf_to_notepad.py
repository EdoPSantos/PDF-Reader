import os
import pdfplumber

input_folder = "../pdfs"
output_folder = "notepads"
os.makedirs(output_folder, exist_ok=True)

# Define a grelha de colunas (em pontos PDF)
colunas_x = [0, 20, 40, 60, 80, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600]

def extract_with_grid(page, colunas_x):
    words = page.extract_words()
    if not words:
        return []

    lines = {}
    for word in words:
        y = round(word["top"])
        x = word["x0"]

        # Determina a coluna da grelha
        column_idx = None
        for i in range(len(colunas_x) - 1):
            if colunas_x[i] <= x < colunas_x[i + 1]:
                column_idx = i
                break

        if column_idx is not None:
            if y not in lines:
                lines[y] = [""] * (len(colunas_x) - 1)
            lines[y][column_idx] += word["text"] + " "

    # Ordena por posição vertical
    ordered_lines = [lines[y] for y in sorted(lines)]
    return ordered_lines

def pdf_to_txt(pdf_path, txt_path):
    all_lines = []
    print(f"A processar: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            grid = extract_with_grid(page, colunas_x)
            if grid:
                all_lines.append(f"\n--- Página {page_num} ---")
                for row in grid:
                    line = "\t".join(cell.strip() for cell in row if cell.strip())
                    all_lines.append(line)

    if all_lines:
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(all_lines))
        print(f"Criado: {txt_path}")
    else:
        print(f"Nenhum conteúdo relevante encontrado: {pdf_path}")

# === Loop pelos PDFs ===
for filename in os.listdir(input_folder):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(input_folder, filename)
        txt_name = os.path.splitext(filename)[0] + ".txt"
        txt_path = os.path.join(output_folder, txt_name)

        if os.path.exists(txt_path):
            print(f"Ignorado (já existe): {txt_name}")
            continue

        pdf_to_txt(pdf_path, txt_path)
