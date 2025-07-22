import sys
import fitz
import json

def extract_words_with_rotation(page):
    word_list = []
    blocks = page.get_text("dict")["blocks"]
    for block in blocks:
        if block.get("type") != 0:
            continue  # Ignora imagens
        for line in block["lines"]:
            for span in line["spans"]:
                angle = span.get("angle", 0)
                width = span["bbox"][2] - span["bbox"][0]
                height = span["bbox"][3] - span["bbox"][1]
                if height > 3 * width or abs(angle) > 5:
                    continue  # Ignora possível texto vertical
                word_list.append({
                    "text": span["text"],
                    "x": span["bbox"][0],
                    "y": span["bbox"][1],
                    "width": span["bbox"][2] - span["bbox"][0],
                    "height": span["bbox"][3] - span["bbox"][1],
                    "font": span.get("font"),
                    "size": span.get("size"),
                    "rotation": angle
                })
    return word_list

def process_pdf(pdf_path, output_path):
    doc = fitz.open(pdf_path)
    pages = []
    for page_number, page in enumerate(doc, start=1):
        words = extract_words_with_rotation(page)
        pages.append({
            "page": page_number,
            "words": words
        })
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(pages, f, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python pdf_to_json.py <pdf_path> <json_path>")
        sys.exit(1)
    pdf_path = sys.argv[1]
    json_path = sys.argv[2]
    process_pdf(pdf_path, json_path)
