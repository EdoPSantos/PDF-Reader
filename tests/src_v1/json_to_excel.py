import sys
import os
import json
import pandas as pd

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python json_to_excel.py <json_path> <excel_path>")
        sys.exit(1)

    json_path = sys.argv[1]
    output_excel = sys.argv[2]

    # Parâmetros da grelha
    row_height = 15
    col_width = 5
    x_scale = 4  # ajusta a escala horizontal
    y_scale = 1  # ajusta a escala vertical

    # Lê o ficheiro JSON recebido
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

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
