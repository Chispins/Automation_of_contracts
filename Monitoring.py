import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import docx
import os
import shutil
# Codigos Propios
from Portada import create_melipilla_document
from Formated_Base_PEP8 import main as main_bases
import pandas as pd
from docxtpl import DocxTemplate
from Jinja_2 import context_for_template1
from Jinja_2 import context_for_template2
from Formated_Base_PEP8 import main as main_bases


class MyHandler(FileSystemEventHandler):
    def __init__(self, root_path):
        self.root_path = os.path.abspath(root_path)  # Carpeta raíz absoluta

    def on_modified(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx",
                                                    "base_automatizada"]:
                return

            # Añadir esta verificación para la modificación del Excel
            if os.path.basename(event.src_path) == "Libro1.xlsx":
                wd = os.path.dirname(event.src_path)
                rendering_base(wd)  # Pasar el directorio de trabajo

            print(f"File modified: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            create_file(wd, self.root_path)

    def on_moved(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) == "Libro1.xlsx":
                rendering_base()

            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "base_automatizada"] or \
                    os.path.basename(event.dest_path) in ["portada_melipilla_base.docx",
                                                          "portada_melipilla_contrato.docx", "base_automatizada"]:
                return
            print(f"File moved: {event.src_path} to {event.dest_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            create_file(wd, self.root_path)
        else:
            if any(name in event.src_path or name in event.dest_path for name in
                   ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx"]):
                return
            print(f"Directory moved: {event.src_path} to {event.dest_path}")
            wd = event.src_path if os.path.exists(event.src_path) else os.path.dirname(event.dest_path)
            print(wd)
            create_file(wd, self.root_path)

    def on_created(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "base_automatizada"]:
                return
            print(f"File created: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            create_file(wd, self.root_path)
        else:
            if any(name in event.src_path for name in
                   ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx"]):
                return
            print(f"Directory created: {event.src_path}")
            wd = event.src_path
            print(wd)
            create_file(wd, self.root_path)

    def on_deleted(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "base_automatizada"]:
                return
            print(f"File deleted: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            create_file(wd, self.root_path)
        else:
            if any(name in event.src_path for name in
                   ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx"]):
                return
            print(f"Directory deleted: {event.src_path}")
            wd = event.src_path
            print(wd)
            create_file(wd, self.root_path)


def MonitoringDirectories(directories):
    handler = MyHandler(directories[0])  # Pasar la carpeta raíz al handler
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


def create_file(wd, root_path):
    try:
        # Normalizar las rutas para comparación
        wd_abs = os.path.abspath(wd)
        # Verificar si wd es la carpeta raíz
        if wd_abs == root_path:
            print(f"Event in root directory {wd}, skipping file creation.")
            return
        # Verificar si el directorio existe
        if not os.path.exists(wd):
            print(f"Directory {wd} does not exist, cannot create file.")
            return

        # Crear las Portadas
        try:
            create_melipilla_document(archivo="base", wd=wd)
            print(f"Documento 'base' creado en {wd}")
        except Exception as e:
            print(f"Error creando documento 'base' en {wd}: {e}")

        try:
            create_melipilla_document(archivo="contrato", wd=wd)
            print(f"Documento 'contrato' creado en {wd}")
        except Exception as e:
            print(f"Error creando documento 'contrato' en {wd}: {e}")

        # Crear un archivo adicional (si es necesario)
        if os.path.exists("base_automatizada"):
            print(f"Archivo 'base_automatizada' ya existe en {wd}, omitiendo creación.")
        else:
            try:
                main_bases(archivo="base", wd=wd, monitoring=True)
            except Exception as e:
                print(f"Error creando archivo 'base_automatizada' en {wd}: {e}")


        """filename = "name.docx"
        output_path = os.path.join(wd, filename)
        if os.path.exists(output_path):
            print(f"File {filename} already exists in {wd}, skipping creation.")
        else:
            try:
                doc = docx.Document()
                doc.add_heading("hola", level=1)
                doc.save(output_path)
                print(f"File {filename} created in {wd}")
            except Exception as e:
                print(f"Error creando archivo adicional {filename} en {wd}: {e}")
"""
        # Copiar el archivo Excel
        source_file = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\NO_MODIFICAR\Libro1.xlsx"
        destination_file = os.path.join(wd, "Libro1.xlsx")
        if os.path.exists(source_file):
            if os.path.exists(destination_file):
                print(f"Archivo Excel 'Libro1.xlsx' ya existe en {wd}, omitiendo copia.")
            else:
                try:
                    shutil.copy(source_file, destination_file)
                    print(f"Archivo Excel 'Libro1.xlsx' copiado a {wd}")
                except Exception as e:
                    print(f"Error copiando archivo Excel a {wd}: {e}")
        else:
            print(f"Error: El archivo de origen '{source_file}' no existe o no es accesible.")
    except Exception as e:
        print(f"Error general creando archivos en {wd}: {e}")


def rendering_base(wd=None):
    """Renderiza los documentos basados en datos de Excel

    Args:
        wd (str, optional): Directorio de trabajo. Si es None, se usa el directorio actual.
    """
    # Si no se proporciona un directorio de trabajo, usar el actual
    if wd is None:
        wd = os.getcwd()

    # Rutas absolutas a los archivos
    excel_path = os.path.join(wd, "Libro1.xlsx")
    base_path = os.path.join(wd, "base_automatizada.docx")
    contrato_path = os.path.join(wd, "contrato_automatizado.docx")

    # Verificar si los archivos existen
    if not os.path.exists(excel_path):
        print(f"Error: No se encuentra el archivo Excel en {excel_path}")
        return

    # Leer datos del Excel
    try:
        data_1 = pd.read_excel(excel_path, sheet_name="Datos_Base")
        triger_1 = data_1.iloc[1, 3]

        data_2 = pd.read_excel(excel_path, sheet_name="Datos_Contrato_P2")
        triger_2 = data_2.iloc[1, 3]

        # Renderizar el documento base
        base_rendered_path = os.path.join(wd, "base_automatizada_rendered.docx")
        if not os.path.exists(base_rendered_path) and triger_1 == "si" and os.path.exists(base_path):
            print(f"Renderizando documento base en {base_rendered_path}")
            doc = DocxTemplate(base_path)
            doc.render(context_for_template1)
            doc.save(base_rendered_path)
            print(f"✅ Documento base renderizado guardado como: {base_rendered_path}")

        # Renderizar el documento de contrato
        contrato_rendered_path = os.path.join(wd, "contrato_automatizado_rendered.docx")
        if not os.path.exists(contrato_rendered_path) and triger_2 == "si" and os.path.exists(contrato_path):
            print(f"Renderizando documento de contrato en {contrato_rendered_path}")
            doc = DocxTemplate(contrato_path)
            doc.render(context_for_template2)
            doc.save(contrato_rendered_path)
            print(f"✅ Documento de contrato renderizado guardado como: {contrato_rendered_path}")

    except Exception as e:
        print(f"Error al renderizar los documentos: {str(e)}")


if __name__ == "__main__":
    path = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\Licitaciones_Testing"
    MonitoringDirectories([path])
