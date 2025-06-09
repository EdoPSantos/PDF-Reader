import os
import json
import pandas as pd

# Caminho da pasta com os ficheiros JSON (altera conforme necessário)
json_folder = "jsons"
output_folder = "excels_visual"

# Cria pasta de saída se não existir
os.makedirs(output_folder, exist_ok=True)

# Parâmetros da grelha
row_height = 15
col_width = 5
x_scale = 4  # ajusta a escala horizontal
y_scale = 1  # ajusta a escala vertical

# Listar ficheiros JSON
json_files = [f for f in os.listdir(json_folder) if f.lower().endswith(".json")]

for json_file in json_files:
    path = os.path.join(json_folder, json_file)
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Criar Excel para este ficheiro
    output_excel = os.path.join(output_folder, f"{os.path.splitext(json_file)[0]}.xlsx")
    writer = pd.ExcelWriter(output_excel, engine='xlsxwriter')

    for page in data:
        page_num = page.get("page", 1)
        words = page.get("words", [])

        grid = {}
        for word in words:
            col = int(word["x"] // (col_width * x_scale))
            row = int(word["y"] // (row_height * y_scale))
            if (row, col) in grid:
                grid[(row, col)] += " " + word["text"]
            else:
                grid[(row, col)] = word["text"]

        max_row = max((r for r, _ in grid.keys()), default=0) + 1
        max_col = max((c for _, c in grid.keys()), default=0) + 1

        df = pd.DataFrame(index=range(max_row), columns=range(max_col))
        for (row, col), text in grid.items():
            df.iat[row, col] = text

        sheet_name = f"pagina_{page_num}"
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    writer.close()
    print(f"Ficheiro Excel criado: {output_excel}")