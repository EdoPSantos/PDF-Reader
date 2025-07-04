import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import sys

PDF_INPUT_FOLDER = os.path.abspath("../../pdfs")
JSON_OUTPUT_FOLDER = os.path.abspath("../jsons")
EXCEL_OUTPUT_FOLDER = os.path.abspath("../excels_visual")

python_executable = sys.executable  # Caminho do Python do ambiente ativo

def pdf_to_json(pdf_path, json_path):
    subprocess.run([python_executable, "pdf_to_json.py", pdf_path, json_path], cwd=os.path.dirname(__file__))

def json_to_excel(json_path, excel_path):
    subprocess.run([python_executable, "json_to_excel.py", json_path, excel_path], cwd=os.path.dirname(__file__))

def process_excel(excel_path):
    subprocess.run([python_executable, "multi_reader.py", excel_path], cwd=os.path.dirname(__file__))

# Dicionário para evitar processamentos repetidos em pouco tempo
recently_processed = {}

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return

        base = os.path.splitext(os.path.basename(event.src_path))[0]
        now = time.time()
        last = recently_processed.get(base, 0)
        if now - last < 10:
            # Ignora se já foi processado há menos de 10 segundos
            return
        recently_processed[base] = now

        print(f"Novo PDF detectado: {event.src_path}")

        # Espera um pouco para garantir que o ficheiro foi completamente copiado
        time.sleep(1)

        json_path = os.path.join(JSON_OUTPUT_FOLDER, f"{base}.json")
        excel_path = os.path.join(EXCEL_OUTPUT_FOLDER, f"{base}.xlsx")
        pdf_to_json(event.src_path, json_path)
        json_to_excel(json_path, excel_path)
        process_excel(excel_path)

if __name__ == "__main__":
    # Caminho absoluto da pasta a monitorizar
    watch_path = PDF_INPUT_FOLDER
    os.makedirs(JSON_OUTPUT_FOLDER, exist_ok=True)
    os.makedirs(EXCEL_OUTPUT_FOLDER, exist_ok=True)
    event_handler = PDFHandler()
    observer = Observer()
    observer.schedule(event_handler, watch_path, recursive=False)
    observer.start()
    print("A monitorizar nova entrada de PDFs em:", watch_path)
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
