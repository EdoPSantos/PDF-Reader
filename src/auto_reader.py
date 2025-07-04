import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess
import sys

PDF_INPUT_FOLDER = os.path.abspath("../pdfs")
JSON_OUTPUT_FOLDER = os.path.abspath("../jsons")

python_executable = sys.executable

def pdf_to_json(pdf_path, json_path):
    subprocess.run([python_executable, "pdf_to_json.py", pdf_path, json_path], cwd=os.path.dirname(__file__))

def process_json(json_path):
    subprocess.run([python_executable, "json_reader.py", json_path], cwd=os.path.dirname(__file__))

processed_files = set()

def is_file_ready(filepath):
    try:
        with open(filepath, "rb"):
            return True
    except Exception:
        return False

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        self.handle_event(event)

    def on_modified(self, event):
        self.handle_event(event)

    def on_deleted(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return
        base = os.path.splitext(os.path.basename(event.src_path))[0]
        processed_files.discard(base)

    def handle_event(self, event):
        if event.is_directory or not event.src_path.lower().endswith(".pdf"):
            return
        base = os.path.splitext(os.path.basename(event.src_path))[0]
        if base in processed_files:
            return

        base = os.path.splitext(os.path.basename(event.src_path))[0]
        if base in processed_files:
            return

        # Espera até o ficheiro estar pronto para ser lido
        for _ in range(10):
            if is_file_ready(event.src_path):
                break
            time.sleep(0.5)
        else:
            print(f"Ficheiro {event.src_path} não está pronto para ser processado.")
            return

        print(f"Novo PDF detectado: {event.src_path}")

        json_path = os.path.join(JSON_OUTPUT_FOLDER, f"{base}.json")
        pdf_to_json(event.src_path, json_path)
        process_json(json_path)
        processed_files.add(base)

if __name__ == "__main__":
    watch_path = PDF_INPUT_FOLDER
    os.makedirs(JSON_OUTPUT_FOLDER, exist_ok=True)
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