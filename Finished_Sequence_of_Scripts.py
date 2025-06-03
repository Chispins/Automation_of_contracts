import time
import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import docx
import pandas as pd
from docxtpl import DocxTemplate
# Importaciones de códigos propios
from Portada import create_melipilla_document
from Formated_Base_PEP8 import main as main_bases
from Formated_Contrats_PEP8_ignore import main as main_contracts
from Jinja_2 import generate_contexts


class MyHandler(FileSystemEventHandler):
    def __init__(self, root_path):
        self.root_path = os.path.abspath(root_path)  # Carpeta raíz absoluta
        self.processed_events = set()  # Para evitar procesar el mismo evento múltiples veces
        #self.files_list = ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "plantilla_original.docx", "Libro1.xlsx","base_automatizada", "base_automatizada_rendered.docx", "plantilla_contrato.docx"]

    def on_modified(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "plantilla_original.docx"]:
                return

            # Añadir esta verificación para la modificación del Excel
            if os.path.basename(event.src_path) == "Libro1.xlsx":
                wd = os.path.dirname(event.src_path)
                self.rendering_base(wd)  # Pasar el directorio de trabajo

            print(f"File modified: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            self.create_files_in_directory(wd)

    def on_created(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "plantilla_original.docx"]:
                return
            print(f"File created: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            self.create_files_in_directory(wd)
        else:
            if any(name in event.src_path for name in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx"]):
                return
            print(f"Directory created: {event.src_path}")
            wd = event.src_path
            print(wd)
            self.create_files_in_directory(wd)

    def on_moved(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "plantilla_original.docx"] or \
                    os.path.basename(event.dest_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "plantilla_original.docx"]:
                return
            print(f"File moved: {event.src_path} to {event.dest_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            self.create_files_in_directory(wd)
        else:
            if any(name in event.src_path or name in event.dest_path for name in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx"]):
                return
            print(f"Directory moved: {event.src_path} to {event.dest_path}")
            wd = event.src_path if os.path.exists(event.src_path) else os.path.dirname(event.dest_path)
            print(wd)
            self.create_files_in_directory(wd)

    def on_deleted(self, event):
        if not event.is_directory:
            if os.path.basename(event.src_path) in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx", "plantilla_original.docx"]:
                return
            print(f"File deleted: {event.src_path}")
            wd = os.path.dirname(event.src_path)
            print(wd)
            self.create_files_in_directory(wd)
        else:
            if any(name in event.src_path for name in ["portada_melipilla_base.docx", "portada_melipilla_contrato.docx"]):
                return
            print(f"Directory deleted: {event.src_path}")
            wd = event.src_path
            print(wd)
            self.create_files_in_directory(wd)

    def create_files_in_directory(self, wd):
        """Crear archivos y copiar recursos en el directorio especificado."""
        try:
            if not os.path.exists(wd):
                print(f"Directorio {wd} no existe, no se pueden crear archivos.")
                return

            # Crear portadas
            for doc_type in ["base", "contrato"]:
                try:
                    create_melipilla_document(archivo=doc_type, wd=wd)
                    print(f"Portada del documento '{doc_type}' creada en {wd}")
                except Exception as e:
                    print(f"Error creando portada del documento '{doc_type}' en {wd}: {e}")

            # Crear archivo automatizado base si no existe
            if os.path.exists(os.path.join(wd, "plantilla_original.docx")):  # Cambiar el nombre del archivo
                print(f"Archivo 'plantilla_original.docx' ya existe en {wd}, omitiendo creación.")
            else:
                try:
                    main_bases()
                    print(f"Archivo 'plantilla_original.docx' creado en {wd}")
                except Exception as e:
                    print(f"Error creando archivo 'plantilla_original.docx' en {wd}: {e}")

            # Copiar archivo Excel si no existe
            source_file = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\NO_MODIFICAR\Libro1.xlsx"
            destination_file = os.path.join(wd, "Libro1.xlsx")
            if os.path.exists(destination_file):
                print(f"Archivo Excel 'Libro1.xlsx' ya existe en {wd}, omitiendo copia.")
            else:
                try:
                    shutil.copy(source_file, destination_file)
                    print(f"Archivo Excel 'Libro1.xlsx' copiado a {wd}")
                except Exception as e:
                    print(f"Error copiando archivo Excel a {wd}: {e}")

            source_file_docx = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\NO_MODIFICAR\plantilla_original.docx"
            destination_file_docx = os.path.join(wd, "plantilla_original.docx")  # Cambiar el nombre del archivo
            if os.path.exists(destination_file_docx):
                print(f"Archivo 'plantilla_original.docx' ya existe en {wd}, omitiendo copia.")
            else:
                try:
                    shutil.copy(source_file_docx, destination_file_docx)
                    print(f"Archivo 'plantilla_original.docx' copiado a {wd}")
                except Exception as e:
                    print(f"Error copiando archivo 'plantilla_original.docx' a {wd}: {e}")

        except Exception as e:
            print(f"Error general creando archivos en {wd}: {e}")

    def rendering_base(self, wd):
        """Renderiza documentos basados en datos de Excel según los triggers."""
        try:
            excel_path = os.path.join(wd, "Libro1.xlsx")
            if not os.path.exists(excel_path):
                print(f"Error: No se encuentra el archivo Excel en {excel_path}")
                return

            # Leer datos del Excel para verificar los triggers
            data_1 = pd.read_excel(excel_path, sheet_name="Datos_Base")
            triger_1 = data_1.iloc[1, 3] if len(data_1) > 1 and len(data_1.columns) > 3 else "no"

            data_2 = pd.read_excel(excel_path, sheet_name="Datos_Contrato_P2")
            triger_2 = data_2.iloc[1, 3] if len(data_2) > 1 and len(data_2.columns) > 3 else "no"
            # Generar contextos desde Jinja_2.py
            context_for_template1, context_for_template2 = generate_contexts(wd)
            if context_for_template1 is None or context_for_template2 is None:
                print(f"Error: No se pudieron generar contextos para renderizado en {wd}")
                return

            # Renderizar documento base si triger_1 es "si"
            base_path = os.path.join(wd, "plantilla_original.docx")  # Cambiar el nombre del archivo
            base_rendered_path = os.path.join(wd, "plantilla_original_rendered.docx")  # Cambiar el nombre del archivo
            if triger_1 == "si" and os.path.exists(base_path) and not os.path.exists(base_rendered_path):
                doc = DocxTemplate(base_path)
                doc.render(context_for_template1)
                doc.save(base_rendered_path)
                print(f"✅ Documento base completo renderizado guardado como: {base_rendered_path}")
            elif triger_1 != "si":
                print(f"Renderizado de base omitido en {wd}: triger_1 no es 'si' (triger_1={triger_1})")
            elif not os.path.exists(base_path):
                print(f"Error: Plantilla base no encontrada en {base_path}")
            elif os.path.exists(base_rendered_path):
                print(f"Archivo base completo renderizado ya existe en {base_rendered_path}, omitiendo renderizado.")

            # Crear y renderizar documento contrato si triger_2 es "si"
            contrato_path = os.path.join(wd, "contrato_automatizado_tablas.docx")
            contrato_rendered_path = os.path.join(wd, "contrato_automatizado_tablas_rendered.docx")
            if triger_2 == "si":
                # Primero, crear la plantilla de contrato si no existe
                if not os.path.exists(contrato_path):
                    try:
                        success = main_contracts(wd=wd, monitoring=True)
                        if success:
                            print(f"Archivo 'contrato_automatizado_tablas.docx' creado en {wd}")
                        else:
                            print(f"Fallo al crear 'contrato_automatizado_tablas.docx' en {wd}")
                    except Exception as e:
                        print(f"Error creando archivo 'contrato_automatizado_tablas.docx' en {wd}: {e}")

                # Luego, renderizar el contrato si la plantilla existe y el archivo renderizado no
                if os.path.exists(contrato_path) and not os.path.exists(contrato_rendered_path):
                    doc = DocxTemplate(contrato_path)
                    doc.render(context_for_template2)
                    doc.save(contrato_rendered_path)
                    print(f"✅ Documento contrato completo renderizado guardado como: {contrato_rendered_path}")
                elif not os.path.exists(contrato_path):
                    print(f"Error: Plantilla contrato no encontrada en {contrato_path}")
                elif os.path.exists(contrato_rendered_path):
                    print(f"Archivo contrato completo renderizado ya existe en {contrato_rendered_path}, omitiendo renderizado.")
            else:
                print(f"Renderizado de contrato omitido en {wd}: triger_2 no es 'si' (triger_2={triger_2})")
        except Exception as e:
            print(f"Error al renderizar documentos en {wd}: {e}")

#template_name_2 = context_for_tempate_"

def monitor_directories(directories):
    """Iniciar el monitoreo de los directorios especificados."""
    if not directories:
        print("No se proporcionaron directorios para monitorear.")
        return

    handler = MyHandler(directories[0])  # Usar el primer directorio como raíz
    observer = Observer()
    for directory in directories:
        if os.path.exists(directory):
            observer.schedule(handler, directory, recursive=True)
            print(f"Monitoreando directorio: {directory}")
        else:
            print(f"Directorio no encontrado: {directory}")

    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("Monitoreo detenido por el usuario.")
    observer.join()


if __name__ == "__main__":
    #path = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\Licitaciones_Testing"
    path = r"C:\Users\Thinkpad\PycharmProjects\Automation_of_contracts\Files\Monitored"
    # Ruta_Alternativa
    path_files = r"C:\Users\Thinkpad\PycharmProjects\Automation_of_contracts\Files"

    monitor_directories([path])
