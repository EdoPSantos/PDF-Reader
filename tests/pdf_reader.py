import pdfplumber
import os

folder = "pdfs"

# Lista os PDFs disponíveis
pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]

if not pdf_files:
    print("Nenhum PDF encontrado na pasta.")
    exit()

# Mostra os PDFs ao utilizador
print("Escolha um ficheiro PDF:")
for i, f in enumerate(pdf_files):
    print(f"{i+1}. {f}")

# Escolha do utilizador
choice = input("Número do ficheiro: ").strip()
if not choice.isdigit() or int(choice) < 1 or int(choice) > len(pdf_files):
    print("Opção inválida.")
    exit()

# Caminho para o ficheiro escolhido
file_path = os.path.join(folder, pdf_files[int(choice) - 1])

# Abrir e mostrar o conteúdo com pdfplumber
with pdfplumber.open(file_path) as pdf:
    print("\n--- Conteúdo do PDF ---\n")
    for page in pdf.pages:
        text = page.extract_text()
        print(text)
        print("\n------------------------\n")
