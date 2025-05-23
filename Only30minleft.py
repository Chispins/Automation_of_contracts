import os
import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd
import psutil
import shutil
# Definimos el directorio a monitorear
directorio = r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\Licitaciones_Testing"
import docx
doc = docx.Document()
cwd = os.getcwd()

class Handler(FileSystemEventHandler):
    def on_modified(self, event):
        # Verificamos si el evento es un archivo y no un directorio
        if event.is_directory:
            return

        file_name = os.path.basename(event.src_path)
        # Filtramos archivos temporales y ocultos
        if not file_name.startswith("~$") and not file_name.endswith(".tmp"):
            if file_name == "Libro1.xlsx":
                try:
                    # Leemos el archivo Excel
                    data = pd.read_excel(event.src_path, sheet_name = "Datos_Base")
                    data_contrato = pd.read_excel(event.src_path, sheet_name = "Datos_Contrato_P1")
                    # Verificamos si existe la fila 3, columna "si"
                    # Nota: pandas usa índices base-0, así que fila 3 es índice 2
                    # Asumiendo que buscas en la columna 1 (índice 0)
                    directorio_carpeta_monitoreada = os.path.dirname(event.src_path)

                    if data.iloc[1,3] == "si":
                        if "base_automatizada.docx" not in os.listdir(directorio_carpeta_monitoreada):
                            print("Base automatizada inicada.")
                            generar_base()
                            render_base()
                        else:
                            print("Base automatizada ya existe.")
                    elif (data_contrato.iloc[1,3] == "si"):
                        if "contrato_automatizado" not in os.listdir(directorio_carpeta_monitoreada):
                            print("Contrato automatizado iniciado.")
                            generar_contrato()
                            render_contrato()
                        else:
                            print("Contrato automatizado ya existe.")

                except Exception as e:
                    print(f"Error al procesar {file_name}: {e}")

    def on_created(self, event):
        # Usamos directamente el event.src_path que ya es una ruta absoluta
        directorio_carpeta_creada = event.src_path

        print(f"Carpeta monitoreada es {directorio_carpeta_creada}")

        if event.is_directory:
            try:
                # Creamos documentos separados para cada caso
                doc_base = docx.Document()
                doc_contrato = docx.Document()
                name_file_origin = "Libro1.xlsx"
                file_origin = os.path.join(r"\\10.5.130.24\Abastecimiento\Compartido Abastecimiento\Otros\NO_MODIFICAR", name_file_origin)
                destination = os.path.join(directorio_carpeta_creada, name_file_origin)
                if os.path.exists(destination):
                    print(f"El archivo {name_file_origin} ya existe en {destination}.")
                else:
                    shutil.copy2(file_origin, destination)
                    print(f"Archivo copiado a {destination}")


                if "portada_melipilla_base" not in os.listdir(directorio_carpeta_creada):
                    print("Creando portada melipilla base...")
                    ruta_completa = os.path.join(directorio_carpeta_creada, "contrato_base.docx")
                    doc_base.save(ruta_completa)
                    print(f"Archivo guardado en: {ruta_completa}")
                    crear_portada_melipilla_base()

                if "portada_melipilla_contrato" not in os.listdir(directorio_carpeta_creada):
                    print("Creando portada melipilla contrato...")
                    ruta_completa = os.path.join(directorio_carpeta_creada, "contrato_portada.docx")
                    doc_contrato.save(ruta_completa)
                    print(f"Archivo guardado en: {ruta_completa}")
                    crear_portada_melipilla_contrato()

            except Exception as e:
                print(f"Error al procesar directorio: {e}")

def generar_base():
    # Aquí iría la lógica para generar la base automatizada
    print("Generando base automatizada...")

def generar_contrato():
    # Aquí iría la lógica para generar el contrato automatizado
    print("Generando contrato automatizado...")

def render_base():
    print("render")

def render_contrato():
    print("render contrato")

def crear_portada_melipilla_base():
    print("Creando portada melipilla base...")

def crear_portada_melipilla_contrato():
    print("Creando portada melipilla contrato...")

def iniciar_monitoreo(path):
    # Nos aseguramos de que el directorio exista
    if not os.path.exists(path):
        print(f"El directorio {path} no existe.")
        return

    # Cambiamos al directorio especificado
    os.chdir(path)

    observer = Observer()
    observer.schedule(Handler(), path=".", recursive=True)
    observer.start()

    print(f"Monitoreando cambios en {path}...")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Monitoreo detenido por el usuario")
        observer.stop()

    observer.join()


# Ejecutamos la función de monitoreo
if __name__ == "__main__":
    iniciar_monitoreo(directorio)