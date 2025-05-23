import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import docx
import os

# Codigos Propios
from Portada import create_melipilla_document







class MyHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) == "qwerty.docx":
                return
            print(f"File modified: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            create_file(wd)

    def on_moved(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) == "qwerty.docx" or os.path.basename(event.dest_path) == "qwerty.docx":
                return
            print(f"File moved: {event.src_path} to {event.dest_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            create_file(wd)
        else:
            if "qwerty.docx" in event.src_path or "qwerty.docx" in event.dest_path:
                return
            print(f"Directory moved: {event.src_path} to {event.dest_path}")
            wd = event.src_path if os.path.exists(event.src_path) else os.path.dirname(event.dest_path)
            print(wd)
            create_file(wd)

    def on_created(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) == "qwerty.docx":
                return
            print(f"File created: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            create_file(wd)
        else:
            if "qwerty.docx" in event.src_path:
                return
            print(f"Directory created: {event.src_path}")
            wd = event.src_path
            print(wd)
            create_file(wd)

    def on_deleted(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) == "qwerty.docx":
                return
            print(f"File deleted: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            create_file(wd)
        else:
            if "qwerty.docx" in event.src_path:
                return
            print(f"Directory deleted: {event.src_path}")
            wd = event.src_path
            print(wd)
            create_file(wd)

def MonitoringDirectories(directories):
    handler = MyHandler()
    observer = Observer()
    for directory in directories:
        observer.schedule(handler, directory, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def create_file(wd):
    try:
        # Construir la ruta absoluta para el archivo
        filename = "name.docx"
        output_path = os.path.join(wd, filename)


        if wd == path:
            return
        # Verificar si el directorio existe
        if not os.path.exists(wd):
            print(f"Directory {wd} does not exist, cannot create file.")
            return
        # Verificar si el archivo ya existe
        if os.path.exists(output_path):
            print(f"File qwerty.docx already exists in {wd}, skipping creation.")
            return

        # Si no existe, crear el archivo
        create_melipilla_document(archivo="base")
        doc = docx.Document()
        doc.add_heading("hola", level=1)
        doc.save(output_path)




        print(f"File qwerty.docx created in {wd}")
    except Exception as e:
        print(f"Error creating file in {wd}: {e}")
    return

if __name__ == "__main__":
    path = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\Licitaciones_Testing"
    MonitoringDirectories([path])
